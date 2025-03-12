# Google Calendar Colorizer

This Google Apps Script automates the process of colorizing calendar events based on various criteria. It's based on the work of Mathias Wagner, as described in [his Medium article](https://mathiasw.medium.com/automate-your-google-calendar-coloring-4e7b15ed5560). Thanks!

## Purpose

The primary goal of this script is to automatically assign colors to calendar events based on specific keywords in the event title, the number of participants, and the event's location. Additionally, it calculates travel times for events that require transit, creating corresponding travel events in your calendar.

## Features

- **Color Coding**: Automatically colorizes events based on:
  - Event title keywords (e.g., "interview", "tentative", etc.)
  - Number of participants (one-on-one vs. group meetings)
  - Valid locations (excluding virtual meeting platforms)
  
- **Travel Time Calculation**: Automatically creates travel events based on the user's home address and the event's location, factoring in travel time and padding for preparation.

## Installation Instructions

1. **Go to Google Apps**:
   - Go to [AppsScript home](https://script.google.com/).

2. **Create a New Project**:
   - Click on the "+" icon to create a new project.

3. **Copy the Scripts**:
   - Copy the contents of `Code.gs` into the script editor. 
   - Create a new file called `appscript.json` and fill it with the contents from here.

4. **Set Up Home Address**:
   - Go to `Project Settings` > `Script properties` and add a new property with the key `HOME_ADDRESS` and your home address as the value. 
   - Tip: create a random event first in Google Calendar, and start to type your address. Once you've found it, copy the auto-completed one to the above property. 

5. **Save and Authorize**:
   - Save the project and authorize the script to access your calendar and location services.

6. **Set Up Triggers**:
   - The script includes triggers that automatically run the `colorizeCalendar` function when events are updated, and on a daily basis for safety. 
   - In the script editor, select the `initializeTriggers` function from the drop down box, and click Run. 
   - Approve necessary permissions if prompted.
   - Verify the created triggers under the Triggers menu.

## Coloring Logic

The script uses the following criteria to determine the color of calendar events. Colors can be changed via the constants at the top of the script.

- **Tentative Events**: Events marked as tentative (indicated by a "?" at the beginning of the title) are colored gray.
- **Interviews**: Events containing the word "interview" in the title are colored pale red.
- **External Events**: Events with a valid location are colored yellow.
- **One-on-One Meetings**: Events with exactly one participant are colored mauve.
- **Group Meetings**: Events with multiple participants are colored blue.
- **Default Color**: Events that do not match any criteria are set to cyan.

## Travel Events

For events that require travel, the script performs the following:

1. **Travel Time Calculation**: It calculates the travel time from the user's home address to the event location using Google Maps.
2. **Travel Events Creation**: The script creates two travel events:
   - One for the time it takes to get to the event (starting from the calculated travel time before the event).
   - Another for the return trip after the event ends. Both travel events are colored yellow.
3. **Transport Mode**: The script used a default transport mode (car) for events that don't have a specified transport mode. Transport mode can be specified by emojis (🚗 for driving and 🚎 for transit) at the beginning of the event title. This transport emoji is removed from the event title after processing.

## Usage, Contributions

This script is provided as is, and you can do whatever you want with it. 

For any issues or contributions, feel free to reach out or submit a pull request.
