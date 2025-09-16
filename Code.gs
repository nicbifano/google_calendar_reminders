function pullCalendarEventsToSheet() {
  const calendar = CalendarApp.getDefaultCalendar();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const existingData = sheet.getDataRange().getValues();

  // Header or new sheet setup
  if (existingData.length === 0) {
    sheet.appendRow(["Event ID", "Title", "Start Time", "End Time", "Guest Emails", "Description", "Reminder Sent?", "Send Email?"]);
  }

  // Map of event keys to existing rows
  const eventMap = new Map();
  for (let i = 1; i < existingData.length; i++) {
    const row = existingData[i];
    const eventKey = row[0] + "|" + new Date(row[2]).toISOString(); // eventId|startTime
    eventMap.set(eventKey, i);
  }

  const now = new Date();
  const endDate = new Date();
  endDate.setDate(now.getDate() + 2); // Only pull 48h ahead if you're only sending 1-hour reminders

  const events = calendar.getEvents(now, endDate);

  events.forEach(event => {
    const id = event.getId();
    const title = event.getTitle();
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const guestEmails = event.getGuestList().map(g => g.getEmail()).join(", ");
    const description = event.getDescription() || "";
    const eventKey = id + "|" + startTime.toISOString();

    // Default logic: if it came from Calendly, mark Send Email? as "Yes"
    const isCalendly = description.toLowerCase().includes("calendly");

    const row = [
      id,
      title,
      startTime,
      endTime,
      guestEmails,
      description,
      "No",                       // Reminder Sent?
      isCalendly ? "Yes" : "No"   // Send Email?
    ];

    if (eventMap.has(eventKey)) {
      const rowIndex = eventMap.get(eventKey) + 1; // Sheet rows are 1-indexed
      const existingRow = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];

      // Keep existing reminder/send flags
      row[6] = existingRow[6]; // Reminder Sent?
      row[7] = existingRow[7]; // Send Email?

      sheet.getRange(rowIndex, 1, 1, 8).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  });

  Logger.log(`Calendar events synced.`);
}

function sendEmailRemindersFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  // Header indexes
  const EVENT_ID_COL = 0;
  const TITLE_COL = 1;
  const START_COL = 2;
  const GUESTS_COL = 4;
  const DESC_COL = 5;
  const REMINDER_COL = 6;
  const REMINDER_SEND = 7;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const startTime = new Date(row[START_COL]);
    const timeToEvent = (startTime - now) / (1000 * 60); // in minutes
    const reminderSent = row[REMINDER_COL];
    const reminderSend = row[REMINDER_SEND];

    if (
      timeToEvent <= 60 &&
      timeToEvent > 30 && // Only send between 30 and 60 minutes in advance
      reminderSent.toString().toLowerCase() !== "yes" && // if already sent, dont send again
      reminderSend.toString().toLocaleLowerCase() == "yes" // check if we want to send
    ) {
      const title = row[TITLE_COL];
      const guestEmails = row[GUESTS_COL];
      const description = row[DESC_COL];

      const recipients = guestEmails
        ? guestEmails.split(",").map(e => e.trim()).filter(Boolean)
        : [];

      const formattedTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "EEE MMM dd, yyyy @ hh:mm a");

      const subject = `📅 Reminder: ${title}`;
      const body = `You have an upcoming meeting:\n\nTitle: ${title}\nTime: ${formattedTime} PST\n\nDetails:\n${description}`;

      recipients.forEach(email => {
        MailApp.sendEmail(email, subject, body);
      });

      sheet.getRange(i + 1, REMINDER_COL + 1).setValue("Yes");
    }
  }

  Logger.log("1-hour reminders sent where applicable.");
}
