# CLAUDE.md - Google Calendar Colorizer Helper

## Commands
- Deploy: Open script.google.com project and click Deploy > New deployment
- Run main: Select `colorizeCalendar` function and click Run button
- Initialize: Run `initializeTriggers()` function to set up script triggers
- Debug: Add `Logger.log()` statements and view logs in Execution log

## Code Style Guidelines
- **Language**: JavaScript (Google Apps Script)
- **Formatting**: ES6 syntax with const/let variable declarations
- **Naming**: camelCase for variables/functions, UPPER_CASE for constants
- **Organization**: Group related constants at top, main function with nested helpers
- **Error Handling**: Use try/catch blocks with Logger.log for errors
- **Comments**: JSDoc-style documentation for functions
- **Patterns**: Early returns in conditional logic, functional approach

## Important Project Structure
- Calendar event processing in main `colorizeCalendar(e)` function
- Configuration via constants at the top of the file
- HOME_ADDRESS stored in Script Properties (not in code)
- Transport mode indicators: 🚗 for driving, 🚎 for transit
- Event color mapping based on title keywords and participant count