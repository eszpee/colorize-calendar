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
const defaultTransport = '🚗'; //transport for events that don't have a mode of transport in their titles
const transportPadding = 15; //how many minutes should be added to travel times to account for extra (eg: going to the car, etc.)

/* 
   Entry for the whole colorizing magic.
*/
function colorizeCalendar(e) {

  Logger.log('Trigger event object:'+ JSON.stringify(e));
  // Extracting the domain of your email, e.g. company.com
  const myOrg     = //CalendarApp.getDefaultCalendar().getName().split("@")[1];
                    ""; // eszpee change: not using this in personal context.

  const pastDays = 1 // looking 1 day back to catch last minute changes
  const futureDays = 7 * 4 // looking 4 weeks into the future
  
  const now       = new Date()
  const startDate = new Date(now.setDate(now.getDate() - pastDays))
  const endDate   = new Date(now.setDate(now.getDate() + futureDays))
                
  const rawEvents = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate)
    
  // Sort by last updated time, most recent first
  const calendarEvents = rawEvents.sort((a, b) => 
    b.getLastUpdated().getTime() - a.getLastUpdated().getTime()
  );
  
  if (e && e.calendarId) {
    // Calendar event trigger - only need to check most recently modified item
    Logger.log('Triggered by Calendar change');
    Logger.log('Calendar ID:'+ e.calendarId);
    
    // The first event should be the most recently modified
    if (calendarEvents.length > 0) {
      const mostRecentlyModified = calendarEvents[0];
      Logger.log('Most recently modified event:'+ JSON.stringify(mostRecentlyModified));
      Logger.log('Modified at:'+ mostRecentlyModified.getLastUpdated());
      if (SKIPCHECK && skipCheck(mostRecentlyModified)) {
        ;
      }
      else {
        colorizeByRegex(mostRecentlyModified, myOrg);
      }
    }
    else { Logger.log("No events found! This is weird."); }
  } 
  else {
    // Non-calendar trigger or manual execution - process all events
    Logger.log('Non-calendar trigger:'+ JSON.stringify(e));
    Logger.log("Calendar Events to process: " + JSON.stringify(calendarEvents).length);
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
          if (DEBUG) { Logger.log("Skipping already colored / declined event:" + event.getTitle()); }
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
      Logger.log("Removing color from non-tentative event: " + eventTitle)
      event.setColor(DEFAULT_EVENT_COLOR)
      // no need to return because this event could need coloring now
    }

    // Check for tentative events
    if (/^\?/.test(eventTitle) && event.getColor() !== TENTATIVE_EVENT_COLOR) {
      Logger.log("Colorizing tentative event found: " + eventTitle)
      event.setColor(TENTATIVE_EVENT_COLOR)
      return
    }

    // Check for interviews
    if (/interview/.test(eventTitle)) {
      Logger.log("Colorizing interview found: " + eventTitle)
      event.setColor(INTERVIEW_COLOR)
      return
    }

    // Check for non-colorized events with a valid location (not starting with "Google" or "Microsoft Teams" that are videoconferencing)
    const location = event.getLocation();
    if (location && !location.startsWith("Google") && !location.startsWith("Microsoft Teams") && !location.includes("http")) {
      Logger.log("New event found with valid location: "+ eventTitle);
      
      if (/^🚗|^🚎/.test(eventTitle)) {
        // let's create an event for travel time
        const t = eventTitle.match(/^🚗|^🚎/)[0];
        Logger.log("Transport is: " + t);
        const travelTime = calculateTravelTime(event.getLocation(), event.getStartTime(), t);
        if (travelTime) {
          Logger.log("Travel time to " + event.getLocation() + ": " + travelTime);

          const travelEventTitle = "Travel (" + t + ")";
          const travelToStart = new Date(event.getStartTime().getTime() - (travelTime + transportPadding) * 60000);
          //TODO: this could be improved to count the time between event and home and not reuse the travelTime calculated TO the event
          const travelBackEnd = new Date(event.getEndTime().getTime() + (travelTime + transportPadding) * 60000);
          const travelTo = CalendarApp.getDefaultCalendar().createEvent(travelEventTitle, travelToStart, event.getStartTime());
          const travelBack = CalendarApp.getDefaultCalendar().createEvent(travelEventTitle, event.getEndTime(), travelBackEnd);
          event.setTitle(event.getTitle().replace(/^🚗|^🚎/, ''));
        }
      }

      Logger.log("Colorizing event with valid location: " + eventTitle)
      event.setColor(EXTERNAL_EVENT_COLOR)

      return
    }

    // Check for events with participants
    const guestList = event.getGuestList();
    if (guestList.length === 1) {
      Logger.log("Colorizing event with one participant: " + eventTitle)
      event.setColor(ONE_ON_ONE_COLOR)
      return
    }
    if (guestList.length > 1) {
      Logger.log("Colorizing event with multiple participants: " + eventTitle)
      event.setColor(GROUP_MEETING_COLOR)
      return
    }

    // No match found, therefore no colorizing - only print if DEBUG is true
    if (DEBUG) { Logger.log("No matching rule for: " + eventTitle); }
  }


  /* Travel time calculator
    Inputs are destination location and means of transport 
    Output in minutes
  */
  function calculateTravelTime(destination,time,transport) {
    const arrivalTime = new Date(time);
    const directions = Maps.newDirectionFinder()
      .setOrigin(HOME_ADDRESS)
      .setDestination(destination)
      .setMode(transports[transport]) 
      .setArrive(arrivalTime)
      .getDirections();
    const route = directions.routes[0];    
    return Math.round(route.legs[0].duration.value / 60);
  }


}

//Initialization, recreates all triggers
function initializeTriggers() {
  // First, delete all existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create the calendar event trigger
  ScriptApp.newTrigger('colorizeCalendar')
    .forUserCalendar(Session.getActiveUser().getEmail())
    .onEventUpdated()
    .create();
    
  // Create the time-based trigger (running once a day)
  ScriptApp.newTrigger('colorizeCalendar')
     .timeBased()
     .everyDays(1)
     .create();

  // Log installed triggers as confirmation
  Logger.log('Triggers initialized successfully:');
  ScriptApp.getProjectTriggers().forEach(trigger => {
    Logger.log(` - Trigger type: ${trigger.getEventType()}`);
  });
}
