/* Written by Mathias Wagner www.linkedin.com/in/mathias-wagner 
   Modified by Péter Szász https://peterszasz.com 
   Based on 
   https://mathiasw.medium.com/automate-your-google-calendar-coloring-4e7b15ed5560
*/


// Global debug setting
// true: print info for every event
// false: print info for only newly colorized event
const DEBUG = true

// Global skip check setting
// true: skip already colored or declined events for better performance
// false: check all events, regardless of their current color or status
const SKIPCHECK = true

// Color constants
const DEFAULT_EVENT_COLOR = CalendarApp.EventColor.CYAN;
const TENTATIVE_EVENT_COLOR = CalendarApp.EventColor.GRAY;
const EXTERNAL_EVENT_COLOR = CalendarApp.EventColor.YELLOW;
const ONE_ON_ONE_COLOR = CalendarApp.EventColor.MAUVE;
const GROUP_MEETING_COLOR = CalendarApp.EventColor.BLUE;
const INTERVIEW_COLOR = CalendarApp.EventColor.PALE_RED;

// Home address for transit time calculations - set it up in Apps Script Settings / Script Properties
const HOME_ADDRESS = PropertiesService.getScriptProperties().getProperty('HOME_ADDRESS');

const transports = {
  '🚗': Maps.DirectionFinder.Mode.DRIVING,
  '🚎': Maps.DirectionFinder.Mode.TRANSIT
};
const transportPadding = 15; //how many minutes shoul be added to travel times to account for extra (eg: going to the car, etc.)

// Logging function
function log(message) {
  if (DEBUG) {
    console.log(message);
  }
}

/* Entry for the whole colorizing magic.
   Select this function when deploying it and assigning a trigger function
*/
function colorizeCalendar() {
 
  const pastDays = 1 // looking 1 day back to catch last minute changes
  const futureDays = 7 * 4 // looking 4 weeks into the future
 
  const now       = new Date()
  const startDate = new Date(now.setDate(now.getDate() - pastDays))
  const endDate   = new Date(now.setDate(now.getDate() + futureDays))
  // Extracting the domain of your email, e.g. company.com
  const myOrg     = //CalendarApp.getDefaultCalendar().getName().split("@")[1];
                    ""; // eszpee change: not using this in personal context.
 
  // Get all calender events within the defined range
  // For now only from the default calendar
  var calendarEvents = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate)

  log("Calendar default org: " + myOrg)

  // Walk through all events, check and colorize
  for (var i=0; i<calendarEvents.length; i++) {
    // Skip for better performance, else go to colorizing below
    if (SKIPCHECK && skipCheck(calendarEvents[i])) {
      continue
    }
    colorizeByRegex(calendarEvents[i], myOrg)
  }
}


/* Performance tweak: skip all events, that do no longer have the DEFAULT color,
   or have been declined already.
   This avoids overriding user settings and doesn't burn regex / string ops
   for allready adjusted event colors.


   @param CalendarEvent
*/
function skipCheck(event) {
    if ((event.getColor() != "" && event.getColor() != TENTATIVE_EVENT_COLOR) || event.getMyStatus() == CalendarApp.GuestStatus.NO) {
        log("Skipping already colored / declined event:" + event.getTitle())
        return true
    }
    return false
}


/* Actual colorizing of events based on Regex matching.
   Makes only sense for frequent stuff you want to auto colorize.
   Order matters for performance! Function exits after first matching color set.
   
   https://developers.google.com/apps-script/reference/calendar/event-color
   Mapping of Google Calendar color names to API color names (Kudos to Jason!):
   https://lukeboyle.com/blog/posts/google-calendar-api-color-id
   @param CalendarEvent
   @param String
*/
function colorizeByRegex(event, myOrg) {
  // Converting to lower case for easier matching.
  // Keep lower case in mind when defining your regex(s) below!
  const eventTitle = event.getTitle().toLowerCase()
 
  // Check for GRAY events and remove color if not tentative anymore
  if (event.getColor() === TENTATIVE_EVENT_COLOR && !/^\?/.test(eventTitle)) {
    log("Removing color from non-tentative event: " + eventTitle)
    event.setColor(DEFAULT_EVENT_COLOR)
    // no need to return because this event could need coloring now
  }

  // Check for tentative events
  if (/^\?/.test(eventTitle) && event.getColor() !== TENTATIVE_EVENT_COLOR) {
    console.log("Colorizing tentative event found: " + eventTitle)
    event.setColor(TENTATIVE_EVENT_COLOR)
    return
  }

  // Check for interviews
  if (/interview/.test(eventTitle)) {
    console.log("Colorizing interview found: " + eventTitle)
    event.setColor(INTERVIEW_COLOR)
    return
  }

  // Check for events with a valid location (not starting with "Google" or "Microsoft Teams" that are videoconferencing)
  const location = event.getLocation();
  if (location && !location.startsWith("Google") && !location.startsWith("Microsoft Teams") && !location.includes("http")) {
    console.log("Event found with valid location: "+ eventTitle);
    if (/^🚗|^🚎/.test(eventTitle)) {
      // let's create an event for travel time
      const t = eventTitle.match(/^🚗|^🚎/)[0];
      console.log("Transport is: " + t);
      const travelTime = calculateTravelTime(event.getLocation(), event.getStartTime(), t);
      if (travelTime) {
        console.log("Travel time to " + event.getLocation() + ": " + travelTime);

        const travelEventTitle = "Travel (" + t + ")";
        const travelToStart = new Date(event.getStartTime().getTime() - (travelTime + transportPadding) * 60000);
        const travelBackEnd = new Date(event.getEndTime().getTime() + (travelTime + transportPadding) * 60000);
        const travelTo = CalendarApp.getDefaultCalendar().createEvent(travelEventTitle, travelToStart, event.getStartTime());
        const travelBack = CalendarApp.getDefaultCalendar().createEvent(travelEventTitle, event.getEndTime(), travelBackEnd);
        event.setTitle(event.getTitle().replace(/^🚗|^🚎/, ''));
      }
    }

    console.log("Colorizing event with valid location: " + eventTitle)
    event.setColor(EXTERNAL_EVENT_COLOR)

    return
  }

  // Check for events with participants
  const guestList = event.getGuestList();
  if (guestList.length === 1) {
    console.log("Colorizing event with one participant: " + eventTitle)
    event.setColor(ONE_ON_ONE_COLOR)
    return
  }
  if (guestList.length > 1) {
    console.log("Colorizing event with multiple participants: " + eventTitle)
    event.setColor(GROUP_MEETING_COLOR)
    return
  }
  
  // No match found, therefore no colorizing
  log("No matching rule for: " + eventTitle)
}


/* Travel time calculator
  Inputs are destination location and means of transport 
  Output in minutes
*/
function calculateTravelTime(destination,time,transport) {
  const directions = Maps.newDirectionFinder()
    .setOrigin(HOME_ADDRESS)
    .setDestination(destination)
    .setMode(transports[transport]) 
    .getDirections();
  const route = directions.routes[0];    
  return Math.round(route.legs[0].duration.value / 60);
}

/* Check participants for external domain other than myOrg. Requires adjustment
   if you have multiple "internal" domains, like company.com and company.de.
   The first external participant match exits this function.
   @param CalendarEvent
   @param String
*/
function checkForExternal(event, myOrg) {
 
  // Create guest list including organizer -> true parameter
  const guestList = event.getGuestList(true)
 
  // Building guest email and domain arrays, not needed by me.
  // Uncomment if you need it
  /*
  var guestEmails = []
  var guestDomains = []


  for (guest of guestList) {
    guestEmails.push(guest.getEmail())
    guestDomains.push(guest.getEmail().split("@")[1])
    if (DEBUG)
    log("Participant emails are: " + guestEmails)
  }
  */


  for (guest of guestList) {


    // get domain of guest and match to my domain
    if (guest.getEmail().split("@")[1] != myOrg) {
     
      log("External domain found: " +
        guest.getEmail().split("@")[1] + " in " + event.getTitle())
      return true
    }
  }
  return false
}