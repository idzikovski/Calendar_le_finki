using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Google.GData.Calendar;
using Google.GData.Client;
using Google.GData.Extensions;

public partial class Calendar_Calendaring_GoogleEvent : System.Web.UI.Page
{
    CalendarService myCalendarService = new CalendarService("FEITPortal");
    Google.GData.Calendar.EventEntry novNastan = new Google.GData.Calendar.EventEntry();
    protected void Page_Load(object sender, EventArgs e)
    {
        
        if (!this.IsPostBack)
        {
           
            string[] uris;
            string povtoruvanje = "";
            mwGoogleEvent.ActiveViewIndex = 0;



            if ((Request.QueryString["token"]!=null)&&Session["sessionToken"]==null)
                Session["sessionToken"] = AuthSubUtil.exchangeForSessionToken(Request.QueryString["token"], null);
            GAuthSubRequestFactory authFactory = new GAuthSubRequestFactory("cl","FEITPortal");
            authFactory.Token =(string)Session["sessionToken"];
            myCalendarService.RequestFactory=authFactory;
            CalendarQuery query = new CalendarQuery();
            query.Uri = new Uri("http://www.google.com/calendar/feeds/default/owncalendars/full");
            try
            {
                CalendarFeed resultFeed = myCalendarService.Query(query);
                foreach (CalendarEntry entry in resultFeed.Entries)
                {
                    uris = entry.Id.Uri.ToString().Split('/');
                    ListItem li = new ListItem();
                    li.Text = entry.Title.Text;
                    li.Value = uris[8];
                    ddlKalendari.Items.Add(li);

                }
            }
            catch (Exception ex)
            {

                hlGcal1.Visible = true;
                hlNazad1.Visible = true;
                lblGreska.Text = "Немате креирано Google Calendar";
                mwGoogleEvent.ActiveViewIndex = 3;
            }
            

            Calendar.COURSE_EVENTS_INSTANCEDataTable nastan = (Calendar.COURSE_EVENTS_INSTANCEDataTable)Session["gNastan"];



            foreach (Calendar.COURSE_EVENTS_INSTANCERow r in nastan)
            {

                CoursesTier CourseTier = new CoursesTier(Convert.ToInt32(Session["COURSE_ID"]));
                novNastan.Content.Language = r.COURSE_EVENT_ID.ToString() ;
                novNastan.Title.Text =r.TITLE;
                novNastan.Content.Content = "Предмет: " + CourseTier.Name + " Опис: " + r.DESCRIPTION + " Тип: " + r.TYPEDESCRIPTION + " ID:" + r.COURSE_EVENT_ID.ToString();
                Where mesto = new Where();
                mesto.ValueString = r.ROOM;
                novNastan.Locations.Add(mesto);
                int recType = Convert.ToInt32(r.RECURRENCE_TYPE);


                DateTime startDate = Convert.ToDateTime(r.STARTDATE);
                DateTime endDate = Convert.ToDateTime(r.ENDDATE);
                DateTime startTime = Convert.ToDateTime(r.STARTIME);
                DateTime endTime = Convert.ToDateTime(r.ENDTIME);
                TimeSpan span = endTime - startTime;

                if (recType != 0)
                {

                    Recurrence rec = new Recurrence();
                    string recData;
                    string dStart = "DTSTART;TZID=" + System.TimeZone.CurrentTimeZone.StandardName + ":" + startDate.ToString("yyyyMMddT") + startTime.AddHours(-1).ToString("HHmm") + "00\r\n";
                    string dEnd = "DTEND;TZID=" + System.TimeZone.CurrentTimeZone.StandardName + ":" + startDate.ToString("yyyyMMddT") + startTime.AddHours(-1 + span.Hours).ToString("HHmm") + "00\r\n";
                    string rRule = "";
                    povtoruvanje = "<b>Повторување:</b> ";

                    switch (recType)
                    {

                        case 1:
                            rRule = "RRULE:FREQ=DAILY;INTERVAL=" + r.RECURRENCE_NUMBER + ";UNTIL="+endDate.ToString("yyyyMMddTHHmm")+"00Z\r\n";
                            povtoruvanje += "Дневно, секој " + r.RECURRENCE_NUMBER.ToString() + " ден <br> </br> <br> </br>";
                            break;
                        case 2:
                            string daysInWeek = "";
                            string denovi = "";

                            if (r.DAYSINWEEK[0] == '1')
                            {
                                daysInWeek += "MO,";
                                denovi += " Понеделник,";
                            }
                            if (r.DAYSINWEEK[1] == '1')
                            {
                                daysInWeek += "TU,";
                                denovi += " Вторник,";
                            }
                            if (r.DAYSINWEEK[2] == '1')
                            {
                                daysInWeek += "WE,";
                                denovi += " Среда,";
                            }
                            if (r.DAYSINWEEK[3] == '1')
                            {
                                daysInWeek += "TH,";
                                denovi += " Четврток,";
                            }
                            if (r.DAYSINWEEK[4] == '1')
                            {
                                daysInWeek += "FR,";
                                denovi += " Петок,";
                            }
                            if (r.DAYSINWEEK[5] == '1')
                            {
                                daysInWeek += "SA,";
                                denovi += " Сабота,";
                            }
                            if (r.DAYSINWEEK[6] == '1')
                            {
                                daysInWeek += "SU,";
                                denovi += " Недела,";
                            }
                            daysInWeek = daysInWeek.Substring(0, daysInWeek.Length - 1);
                            denovi = denovi.Substring(0, denovi.Length - 1);
                            rRule = "RRULE:FREQ=WEEKLY;INTERVAL=" + r.RECURRENCE_NUMBER + ";BYDAY=" + daysInWeek + ";UNTIL=" + endDate.ToString("yyyyMMddTHHmm") + "00Z\r\n";
                            povtoruvanje += "Неделно, секоја " + r.RECURRENCE_NUMBER + " недела и тоа во: " + denovi + " <br> </br> <br> </br>";
                            break;
                        case 3:
                            rRule = "RRULE:FREQ=MONTHLY;INTERVAL=" + r.RECURRENCE_NUMBER + ";UNTIL=" + endDate.ToString("yyyyMMddTHHmm") + "00Z\r\n";
                            povtoruvanje += "Месечно, секој " + r.RECURRENCE_NUMBER + " месец <br> </br> <br> </br>";
                            break;

                    }
                    recData = dStart + dEnd + rRule;
                    rec.Value = recData;
                    novNastan.Recurrence = rec;

                }
                else
                {
                    When vreme = new When();
                    vreme.StartTime = r.STARTIME;
                    vreme.EndTime = r.ENDTIME;
                    novNastan.Times.Add(vreme);
                }


                lblPrikaz.Text += "<b>Наслов: </b>" + r.TITLE + "<br> </br> <br> </br>";
                lblPrikaz.Text += "<b>Опис: </b>" + r.DESCRIPTION + "<br> </br> <br> </br>";
                lblPrikaz.Text += "<b>Просторија: </b>" + r.ROOM + "<br> </br> <br> </br>";
                lblPrikaz.Text += "<b>Тип: </b>" + r.TYPEDESCRIPTION + "<br> </br> <br> </br>";
                lblPrikaz.Text += povtoruvanje;
                lblPrikaz.Text += "<b>Почетен датум: </b>" + startDate.ToShortDateString() + "  <b>Краен датум: </b>" + endDate.ToShortDateString() + "<br> </br> <br> </br>";
                lblPrikaz.Text += "<b>Време: Од </b>" + startTime.ToShortTimeString() + " <b>До</b> " + endTime.ToShortTimeString();



            }
            Session["novNastan"] = novNastan;

            
        }
    }
    protected void ddlNastani_SelectedIndexChanged(object sender, EventArgs e)
    {
        
    }
    protected void btnPonatamu_Click(object sender, EventArgs e)
    {

        if (!nastanotPostoi())
            mwGoogleEvent.ActiveViewIndex = 1;
        else
        {
            btnNazad.Visible = true;
            lblGreska.Text = "Настанот веќе постои во избраниот календар. Избришете го или одберете друг календар";
            mwGoogleEvent.ActiveViewIndex = 3;
        }

    }
    protected bool nastanotPostoi()
    {
        bool postoi = false;
        Google.GData.Calendar.EventEntry novNastan=(Google.GData.Calendar.EventEntry)Session["novNastan"];
        string calendarId = ddlKalendari.SelectedItem.Value;
        EventQuery evQuery = new EventQuery();
        GAuthSubRequestFactory authFactory = new GAuthSubRequestFactory("cl", "FEITPortal");
        authFactory.Token = (string)Session["sessionToken"];
        myCalendarService.RequestFactory = authFactory;
        evQuery.Uri = new Uri("http://www.google.com/calendar/feeds/" + calendarId + "/private/full");
        Google.GData.Calendar.EventFeed evFeed = myCalendarService.Query(evQuery) as Google.GData.Calendar.EventFeed;

        if (evFeed != null)
        {
            foreach (Google.GData.Calendar.EventEntry en in evFeed.Entries)
            {
                string[] newEventId = novNastan.Content.Content.Split(':');
                string[] eventId = en.Content.Content.Split(':');
                int novLength = newEventId.Length;
                int starLength = eventId.Length;
                if (eventId[starLength-1] == newEventId[novLength-1])
                {
                    postoi = true;
                    break;
                }
             
            }
        }
        return postoi;
    }
    protected void btnDodadi_Click(object sender, EventArgs e)
    {

        if (!nastanotPostoi())
        {
            novNastan = (Google.GData.Calendar.EventEntry)Session["novNastan"];
            string calendarId = ddlKalendari.SelectedItem.Value;
            GAuthSubRequestFactory authFactory = new GAuthSubRequestFactory("cl", "FEITPortal");
            authFactory.Token = (string)Session["sessionToken"];
            myCalendarService.RequestFactory = authFactory;

            Uri postUri = new Uri("http://www.google.com/calendar/feeds/" + calendarId + "/private/full");
            AtomEntry insertedEntry = myCalendarService.Insert(postUri, novNastan);
            lblUspesno.Text = "Настанот беше успешно креиран!";
            mwGoogleEvent.ActiveViewIndex = 2;
            
        }
        else
        {
            btnNazad.Visible = true;
            lblGreska.Text = "Настанот веќе постои во избраниот календар. Избришете го или обидете друг календар";
            mwGoogleEvent.ActiveViewIndex = 3;
        }
        
    }
    protected void btnNazad_Click(object sender, EventArgs e)
    {
        mwGoogleEvent.ActiveViewIndex = 0;
    }
    protected void btnOtkazi_Click(object sender, EventArgs e)
    {
        Response.Redirect("http://localhost:1352/FEITPortal/Calendar/Calendaring/weekly.aspx");
    }
}
