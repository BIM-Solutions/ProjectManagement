import * as React from "react";
import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Input,
  Field,
  Textarea,
  Dropdown,
  Option,
} from "@fluentui/react-components";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IProject } from "../../services/ProjectService";
import { Calendar, Views, momentLocalizer } from "react-big-calendar";
import moment from "moment";
import "react-big-calendar/lib/css/react-big-calendar.css";

// Set up localizer for react-big-calendar using moment.js
const localizer = momentLocalizer(moment);

// Props for CalendarTab component
export interface ICalendarTabProps {
  context: WebPartContext;
  project: IProject | undefined;
}

// Calendar event interface
interface ICalendarEvent {
  id: number;
  title: string;
  start: Date;
  end: Date;
  description: string;
  type: string;
  status: string;
}

// Event type and status options
const eventTypes = [
  "Meeting",
  "Deadline",
  "Review",
  "Site Visit",
  "CofW",
  "Other",
];

const eventStatuses = ["Scheduled", "Completed", "Cancelled", "Pending"];

// Styles for the component using Fluent UI's makeStyles
const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    height: "100%",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  calendar: {
    height: "600px",
  },
  form: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingHorizontalM,
  },
  fullWidth: {
    gridColumn: "1 / -1",
  },
});

/**
 * CalendarTab component displays a calendar for project events.
 */
const CalendarTab: React.FC<ICalendarTabProps> = ({ context, project }) => {
  const styles = useStyles();

  // State for calendar events
  const [events, setEvents] = useState<ICalendarEvent[]>([]);
  // State to control event dialog visibility
  const [showEventDialog, setShowEventDialog] = useState(false);
  // State for currently selected event (for editing)
  const [selectedEvent, setSelectedEvent] = useState<ICalendarEvent | null>(
    null
  );
  // State for event form data
  const [formData, setFormData] = useState<Partial<ICalendarEvent>>({});

  /**
   * Loads events for the current project.
   * Replace mockEvents with actual API call in production.
   */
  const loadEvents = async (): Promise<void> => {
    try {
      // Mock data for demonstration
      const mockEvents: ICalendarEvent[] = [
        {
          id: 1,
          title: "Project Review",
          start: new Date(),
          end: new Date(Date.now() + 2 * 60 * 60 * 1000), // 2 hours later
          description: "Monthly project review meeting",
          type: "Meeting",
          status: "Scheduled",
        },
      ];
      setEvents(mockEvents);
    } catch (error) {
      console.error("Error loading events:", error);
    }
  };

  // Load events when project changes
  useEffect(() => {
    if (project) {
      loadEvents().catch(console.error);
    }
  }, [project]);

  /**
   * Handles changes to form fields.
   * @param field - The field being updated
   * @param value - The new value
   */
  const handleInputChange = (
    field: keyof ICalendarEvent,
    value: string | Date
  ): void => {
    setFormData((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  /**
   * Handles saving a new or edited event.
   */
  const handleSave = async (): Promise<void> => {
    try {
      if (selectedEvent) {
        // Update existing event
        const updatedEvents = events.map((event) =>
          event.id === selectedEvent.id ? { ...event, ...formData } : event
        );
        setEvents(updatedEvents);
      } else {
        // Create new event
        const newEvent: ICalendarEvent = {
          id: Math.max(...events.map((e) => e.id), 0) + 1,
          title: formData.title || "",
          start: formData.start as Date,
          end: formData.end as Date,
          description: formData.description || "",
          type: formData.type || "",
          status: formData.status || "",
        };
        setEvents([...events, newEvent]);
      }
      // Reset dialog and form state
      setShowEventDialog(false);
      setSelectedEvent(null);
      setFormData({});
    } catch (error) {
      console.error("Error saving event:", error);
    }
  };

  /**
   * Handles selecting a time slot on the calendar to create a new event.
   * @param param0 - Object containing start and end dates
   */
  const handleSelectSlot = ({
    start,
    end,
  }: {
    start: Date;
    end: Date;
  }): void => {
    setFormData({
      start,
      end,
    });
    setShowEventDialog(true);
  };

  /**
   * Handles selecting an existing event for editing.
   * @param event - The selected calendar event
   */
  const handleSelectEvent = (event: ICalendarEvent): void => {
    setSelectedEvent(event);
    setFormData(event);
    setShowEventDialog(true);
  };

  return (
    <div className={styles.container}>
      {/* Header with title and new event button */}
      <div className={styles.header}>
        <h2>Project Calendar</h2>
        <Button
          appearance="primary"
          onClick={() => {
            setSelectedEvent(null);
            setFormData({});
            setShowEventDialog(true);
          }}
        >
          New Event
        </Button>
      </div>

      {/* Calendar component */}
      <Calendar
        localizer={localizer}
        events={events}
        startAccessor="start"
        endAccessor="end"
        selectable
        onSelectSlot={handleSelectSlot}
        onSelectEvent={handleSelectEvent}
        defaultView={Views.MONTH}
        views={[Views.MONTH, Views.WEEK, Views.DAY]}
        className={styles.calendar}
      />

      {/* Dialog for creating/editing events */}
      <Dialog open={showEventDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              {selectedEvent ? "Edit Event" : "New Event"}
            </DialogTitle>
            <DialogContent>
              <div className={styles.form}>
                {/* Title input */}
                <Field label="Title" required>
                  <Input
                    value={formData.title || ""}
                    onChange={(_, data) =>
                      handleInputChange("title", data.value)
                    }
                  />
                </Field>

                {/* Type dropdown */}
                <Field label="Type">
                  <Dropdown
                    value={formData.type || ""}
                    onOptionSelect={(_, data) =>
                      handleInputChange("type", data.optionValue || "")
                    }
                  >
                    {eventTypes.map((type) => (
                      <Option key={type} value={type}>
                        {type}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                {/* Status dropdown */}
                <Field label="Status">
                  <Dropdown
                    value={formData.status || ""}
                    onOptionSelect={(_, data) =>
                      handleInputChange("status", data.optionValue || "")
                    }
                  >
                    {eventStatuses.map((status) => (
                      <Option key={status} value={status}>
                        {status}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                {/* Start date picker */}
                <Field label="Start Date">
                  <DatePicker
                    value={
                      formData.start ? new Date(formData.start) : undefined
                    }
                    onSelectDate={(date: Date | null | undefined) => {
                      if (date) {
                        handleInputChange("start", date);
                      }
                    }}
                  />
                </Field>

                {/* End date picker */}
                <Field label="End Date">
                  <DatePicker
                    value={formData.end ? new Date(formData.end) : undefined}
                    onSelectDate={(date: Date | null | undefined) => {
                      if (date) {
                        handleInputChange("end", date);
                      }
                    }}
                  />
                </Field>

                {/* Description textarea */}
                <Field label="Description" className={styles.fullWidth}>
                  <Textarea
                    value={formData.description || ""}
                    onChange={(_, data) =>
                      handleInputChange("description", data.value)
                    }
                  />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              {/* Cancel button */}
              <Button
                appearance="secondary"
                onClick={() => setShowEventDialog(false)}
              >
                Cancel
              </Button>
              {/* Save button */}
              <Button appearance="primary" onClick={handleSave}>
                Save
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default CalendarTab;
