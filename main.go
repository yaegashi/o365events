package main

import (
	"bytes"
	"context"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"regexp"
	"sort"
	"strings"
	"time"

	"github.com/tealeg/xlsx/v2"
	"github.com/yaegashi/msgraph.go/jsonx"
	"github.com/yaegashi/msgraph.go/msauth"
	msgraph "github.com/yaegashi/msgraph.go/v1.0"
	"golang.org/x/oauth2"
)

const (
	defaultTenantID       = "common"
	defaultClientID       = "b7dbe94f-2f3a-4b98-a372-a99d0edff196"
	defaultTokenCachePath = "token_cache.json"
	defaultOutputPath     = "events.xlsx"
	dateTimeEventFormat   = "2006-01-02T15:04:05.999999999"
	dateTimeExcelformat   = "2006-01-02 15:04:05"
	dateFlagFormat        = "20060102"
)

var defaultScopes = []string{"offline_access", "User.ReadBasic.All", "Calendars.Read", "Calendars.Read.Shared", "Sites.Read.All", "Files.ReadWrite.All"}

func dump(o interface{}) {
	enc := jsonx.NewEncoder(os.Stdout)
	enc.SetIndent("", "  ")
	enc.Encode(o)
}

type CalEvent struct {
	Start     time.Time
	End       time.Time
	Subject   string
	Location  string
	Organizer string
	Attendees []string
	WebLink   string
}

type CalUser struct {
	DisplayName string
	Events      []*CalEvent
}

type App struct {
	TenantID       string
	ClientID       string
	TokenCachePath string
	Output         string
	Exclude        bool
	Start          time.Time
	End            time.Time
	FlagSet        *flag.FlagSet
	HTTPClient     *http.Client
	GraphClient    *msgraph.GraphServiceRequestBuilder
}

var escapeSheetNameRe = regexp.MustCompile(`[/\\\?\*\[\]]`)

func escapeSheetName(name string) string {
	return escapeSheetNameRe.ReplaceAllString(name, "_")
}

func convertEmailAddressToString(a *msgraph.EmailAddress) string {
	return fmt.Sprintf("%s <%s>", *a.Name, *a.Address)
}

func convertTimeZoneToLocation(tz string) *time.Location {
	loc, err := time.LoadLocation(tz)
	if err != nil {
		loc = time.Local
	}
	return loc
}

func (app *App) fetchCalUsers(ctx context.Context, args []string) ([]*CalUser, error) {
	if len(args) == 0 {
		args = []string{"me"}
	}
	calUsers := []*CalUser{}
	for _, arg := range args {
		log.Printf("I: User %s", arg)
		var (
			user *msgraph.User
			err  error
		)
		if arg == "me" {
			user, err = app.GraphClient.Me().Request().Get(ctx)
			if err != nil {
				log.Printf("E: %s", err)
				continue
			}
		} else {
			user, err = app.GraphClient.Users().ID(arg).Request().Get(ctx)
			if err != nil {
				if res, ok := err.(*msgraph.ErrorResponse); !ok || res.StatusCode() != http.StatusNotFound {
					log.Printf("E: %s", err)
					continue
				}
				req := app.GraphClient.Users().Request()
				req.Filter(fmt.Sprintf("mail eq '%s'", arg))
				users, err := req.Get(ctx)
				if err != nil {
					log.Printf("E: %s", err)
					continue
				}
				if len(users) != 1 {
					log.Printf("E: No unique user found for %s, skipping", arg)
					continue
				}
				user = &users[0]
			}
		}

		log.Printf("I: Fetching events of %s (%s)", *user.UserPrincipalName, *user.ID)
		calEvents := []*CalEvent{}
		req := app.GraphClient.Users().ID(*user.ID).Calendar().CalendarView().Request()
		req.Top(100)
		req.Header().Add("Prefer", `outlook.timezone="Tokyo Standard Time"`)
		req.Query().Add("startDateTime", app.Start.Format(time.RFC3339))
		req.Query().Add("endDateTime", app.End.Format(time.RFC3339))
		oEvents, err := req.Get(ctx)
		if err != nil {
			log.Println(err)
			continue
		}
		log.Printf("I: Got %d events", len(oEvents))

		for _, oEvent := range oEvents {
			organizer := ""
			if oEvent.Organizer != nil {
				organizer = convertEmailAddressToString(oEvent.Organizer.EmailAddress)
			}
			start, _ := time.ParseInLocation(dateTimeEventFormat, *oEvent.Start.DateTime, convertTimeZoneToLocation((*oEvent.Start.TimeZone)))
			end, _ := time.ParseInLocation(dateTimeEventFormat, *oEvent.End.DateTime, convertTimeZoneToLocation(*oEvent.End.TimeZone))
			calEvent := &CalEvent{
				Subject:   *oEvent.Subject,
				Start:     start.In(time.Local),
				End:       end.In(time.Local),
				Location:  *oEvent.Location.DisplayName,
				Organizer: organizer,
				WebLink:   *oEvent.WebLink,
			}
			for _, attendee := range oEvent.Attendees {
				if !app.Exclude || (user.Mail != nil && attendee.EmailAddress.Address != nil && *user.Mail != *attendee.EmailAddress.Address) {
					calEvent.Attendees = append(calEvent.Attendees, convertEmailAddressToString(attendee.EmailAddress))
				}
			}
			calEvents = append(calEvents, calEvent)
		}
		sort.Slice(calEvents, func(i, j int) bool { return calEvents[i].Start.Before(calEvents[j].Start) })
		calUser := &CalUser{
			DisplayName: *user.DisplayName,
			Events:      calEvents,
		}
		calUsers = append(calUsers, calUser)
	}
	return calUsers, nil
}

func (app *App) generateExcelFile(ctx context.Context, calUsers []*CalUser) (*xlsx.File, error) {
	file := xlsx.NewFile()
	for _, calUser := range calUsers {
		sheet, err := file.AddSheet(escapeSheetName(calUser.DisplayName))
		if err != nil {
			return nil, err
		}
		header := []string{"Outlook", "Start", "End", "Subject", "Location", "Organizer", "Attendees"}
		row := sheet.AddRow()
		row.WriteSlice(&header, -1)
		for _, calEvent := range calUser.Events {
			row := sheet.AddRow()
			n := len(calEvent.Attendees)
			if n == 0 {
				n = 1
			}
			row.SetHeight(float64(n*15 + 2))
			c := row.AddCell()
			c.SetHyperlink(calEvent.WebLink, "LINK", "")
			c = row.AddCell()
			c.SetString(calEvent.Start.Local().Format(dateTimeExcelformat))
			c = row.AddCell()
			c.SetString(calEvent.End.Local().Format(dateTimeExcelformat))
			row.AddCell().SetString(calEvent.Subject)
			row.AddCell().SetString(calEvent.Location)
			row.AddCell().SetString(calEvent.Organizer)
			c = row.AddCell()
			c.SetString(strings.Join(calEvent.Attendees, "\r\n"))
			style := xlsx.NewStyle()
			style.Font.Size = 11
			style.Alignment.WrapText = true
			style.Alignment.Vertical = "center"
			c.SetStyle(style)
		}
		sheet.SetColWidth(2, 3, 20)
		sheet.SetColWidth(4, 6, 60)
		sheet.SetColWidth(7, 7, 80)
	}
	return file, nil
}

func (app *App) uploadToSPO(ctx context.Context, b io.Reader) error {
	itemRB, err := app.GraphClient.GetDriveItemByURL(ctx, app.Output)
	if err != nil {
		return err
	}
	req, err := itemRB.Request().NewRequest("PUT", "/content", b)
	if err != nil {
		return err
	}
	req = req.WithContext(ctx)
	res, err := app.HTTPClient.Do(req)
	if err != nil {
		return err
	}
	defer res.Body.Close()
	switch res.StatusCode {
	case http.StatusOK, http.StatusCreated:
		return nil
	default:
		b, _ := ioutil.ReadAll(res.Body)
		return fmt.Errorf("%s: %s", res.Status, string(b))
	}
}

func (app *App) main(args []string) error {
	startString := time.Now().Local().Format(dateFlagFormat)
	endString := ""

	app.FlagSet = flag.NewFlagSet(args[0], flag.ExitOnError)
	app.FlagSet.StringVar(&app.TenantID, "tenant-id", defaultTenantID, "Tenant ID")
	app.FlagSet.StringVar(&app.ClientID, "client-id", defaultClientID, "Client ID")
	app.FlagSet.StringVar(&app.TokenCachePath, "token-cache-path", defaultTokenCachePath, "Token cache path")
	app.FlagSet.StringVar(&app.Output, "output", defaultOutputPath, "Output path")
	app.FlagSet.StringVar(&startString, "start", startString, "Start month (YYYYMM)")
	app.FlagSet.StringVar(&endString, "end", endString, "End month (YYYYMM)")
	app.FlagSet.BoolVar(&app.Exclude, "exclude", false, "Exclude calendar owner from attendees")
	app.FlagSet.Parse(args[1:])

	var err error
	app.Start, err = time.Parse(dateFlagFormat, startString)
	if err != nil {
		return err
	}
	if endString == "" {
		endString = startString
	}
	app.End, err = time.Parse(dateFlagFormat, endString)
	if err != nil {
		return err
	}
	app.End = time.Date(app.End.Year(), app.End.Month(), app.End.Day()+1, 0, 0, 0, 0, time.Local)

	ctx := context.Background()
	m := msauth.NewManager()
	m.LoadFile(app.TokenCachePath)
	ts, err := m.DeviceAuthorizationGrant(ctx, app.TenantID, app.ClientID, defaultScopes, nil)
	if err != nil {
		return err
	}
	m.SaveFile(app.TokenCachePath)

	app.HTTPClient = oauth2.NewClient(ctx, ts)
	app.GraphClient = msgraph.NewClient(app.HTTPClient)

	calUsers, err := app.fetchCalUsers(ctx, app.FlagSet.Args())
	if err != nil {
		return err
	}

	log.Printf("I: Writing to %s", app.Output)
	b := &bytes.Buffer{}
	if strings.HasSuffix(app.Output, ".xlsx") {
		excelFile, err := app.generateExcelFile(ctx, calUsers)
		if err != nil {
			return err
		}
		err = excelFile.Write(b)
		if err != nil {
			return err
		}
	} else if app.Output == "-" || strings.HasSuffix(app.Output, ".json") {
		enc := jsonx.NewEncoder(b)
		enc.SetIndent("", "  ")
		err = enc.Encode(calUsers)
		if err != nil {
			return err
		}
	} else {
		return fmt.Errorf("E: Format unknown for %s", app.Output)
	}

	if strings.HasPrefix(app.Output, "https://") {
		err = app.uploadToSPO(ctx, b)
		if err != nil {
			return err
		}
	} else {
		file := os.Stdout
		if app.Output != "-" {
			file, err = os.Create(app.Output)
			if err != nil {
				return err
			}
			defer file.Close()
		}
		_, err = io.Copy(file, b)
		if err != nil {
			return err
		}
	}

	return nil
}

func main() {
	app := &App{}
	err := app.main(os.Args)
	if err != nil {
		log.Fatalf("E: %s", err)
	}
}
