import * as React from "react";
import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  Text,
  Spinner,
  Tooltip,
  // Dropdown,
  // Option,
  Switch,
  Persona,
  // Dialog,
  // DialogTitle,
  // DialogContent,
  // DialogActions,
  // Button,
  // DialogSurface,
  // DialogBody,
} from "@fluentui/react-components";
// import { DateRangeType } from "@fluentui/react-calendar-compat";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
// import moment from "moment";
import "react-big-calendar/lib/css/react-big-calendar.css";
// import {
//   // Calendar as BigCalendar,
//   momentLocalizer,
//   // Views,
//   // EventProps,
// } from "react-big-calendar";
// import { spfi, SPFI } from "@pnp/sp";
// import { SPFx } from "@pnp/sp/presets/all";
import FullCalendar from "@fullcalendar/react";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import interactionPlugin from "@fullcalendar/interaction";
import listPlugin from "@fullcalendar/list";
import multiMonthPlugin from "@fullcalendar/multimonth";
// import "@fullcalendar/core/index.css";
// import "@fullcalendar/daygrid/index.css";
// import "@fullcalendar/timegrid/index.css";
import {
  EventContentArg,
  EventChangeArg,
  EventClickArg,
  EventApi,
} from "@fullcalendar/core";
import styles from "./CalendarView.module.scss";
import TasksService, { ITask } from "../../services/TasksService";
import { TaskAddEditDialog } from "../../../common/components/Event/TaskAddEditDialog";
// const localizer = momentLocalizer(moment);

const useStyles = makeStyles({
  main: {
    display: "flex",
    flexDirection: "row",
    width: "100%",
    height: "100%",
  },
  calendarContainer: {
    flex: 2,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    minWidth: 0,
  },
  rightPanel: {
    flex: 1,
    backgroundColor: tokens.colorNeutralBackground2,
    padding: tokens.spacingVerticalL,
    minWidth: "320px",
    maxWidth: "400px",
    borderLeft: `1px solid ${tokens.colorNeutralStroke1}`,
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    marginLeft: "20px",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  eventCard: {
    padding: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  calendar: {
    height: "600px",
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "20px",
    borderRadius: "12px",
  },
  eventWrapper: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    width: "100%",
    overflow: "hidden",
  },
  eventTitle: {
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  controls: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    alignItems: "center",
  },
  eventsList: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
});

interface ICalendarEvent {
  id: string;
  subject: string;
  title: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  organizer: { emailAddress: { name: string; address: string } };
  isAllDay: boolean;
  resource?: TaskItem;
}

interface ICalendarViewProps {
  showTeamCalendar: boolean;
  groupId: string;
  context: WebPartContext;
  userDisplayName: string;
  listName: string;
}

interface TaskItem {
  Id: number;
  Title: string;
  DueDate?: string;
  StartDate?: string;
  AssignedTo?: {
    Title: string;
    EMail: string;
  };
  Priority?: string;
  Status?: string;
  ProjectID?: string;
}

type CalendarEventForDisplay = Omit<ICalendarEvent, "start" | "end"> & {
  start: Date;
  end: Date;
  eventType: string;
  ProjectID?: string;
};

type FullCalendarEvent = {
  id: string;
  title: string;
  start: Date;
  end: Date;
  color: string;
  extendedProps?: {
    resource?: TaskItem;
    organizer?: { emailAddress: { name: string; address: string } };
    eventType?: string;
    ProjectID?: string;
  };
};

export const CalendarView: React.FC<ICalendarViewProps> = (props) => {
  const [events, setEvents] = useState<CalendarEventForDisplay[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [includeTasks, setIncludeTasks] = useState(true);
  const [showWeekends, setShowWeekends] = useState(false);
  const [showBusinessHours, setShowBusinessHours] = useState(false);
  const [graphClient, setGraphClient] = useState<MSGraphClientV3 | null>(null);
  const [calendarKey, setCalendarKey] = useState<number>(Date.now());
  const stylesFluent = useStyles();
  const tasksService = TasksService.getInstance(props.context);
  // const [selectedEvent, setSelectedEvent] = useState<FullCalendarEvent | null>(
  //   null
  // );
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedEvent, setSelectedEvent] = useState<EventApi | null>(null);

  useEffect(() => {
    props.context.msGraphClientFactory
      .getClient("3")
      .then((client) => setGraphClient(client))
      .catch(console.error);
  }, [props.context]);

  const mapEvent = (
    event: ICalendarEvent,
    eventType: string
  ): CalendarEventForDisplay => ({
    ...event,
    start: new Date(event.start.dateTime),
    end: new Date(event.end.dateTime),
    eventType: eventType,
  });

  const mapTaskToEvent = (task: TaskItem): CalendarEventForDisplay | null => {
    if (!task.DueDate) return null;
    const start = new Date(task.DueDate);
    // Make it an all-day event (end is next day)
    const end = new Date(start);
    end.setDate(start.getDate() + 1);
    return {
      eventType: "Task",
      id: `task-${task.Id}`,
      subject: task.Title || "",
      title: task.Title || "",
      start: new Date(task.StartDate || task.DueDate || new Date()),
      end: new Date(task.DueDate || new Date()),
      organizer: {
        emailAddress: {
          name: task.AssignedTo?.Title || "",
          address: task.AssignedTo?.EMail || "",
        },
      },
      resource: task,
      isAllDay: true,
      ProjectID: task.ProjectID || "",
    };
  };

  const loadEvents = async (): Promise<void> => {
    try {
      setIsLoading(true);
      let calendarEvents: CalendarEventForDisplay[] = [];
      let spTaskEvents: CalendarEventForDisplay[] = [];

      // Always fetch personal calendar events
      const client =
        graphClient ||
        (await props.context.msGraphClientFactory.getClient("3"));
      if (!graphClient) setGraphClient(client);
      const personalResponse = await client.api("/me/calendar/events").get();
      calendarEvents = personalResponse.value.map((event: ICalendarEvent) => ({
        ...mapEvent(event, "Calender"),
        title: event.subject,
      }));

      // If including tasks, fetch from SharePoint and merge
      if (includeTasks) {
        const tasks = await tasksService.getTasks(
          "9719_ProjectTasks",
          props.userDisplayName
        );
        spTaskEvents = tasks
          .map(mapTaskToEvent)
          .filter((ev): ev is CalendarEventForDisplay => ev !== null);
      }

      setEvents([...calendarEvents, ...spTaskEvents]);
      // setTeamEvents([...calendarEvents, ...spTaskEvents]);
    } catch (err) {
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadEvents().catch(console.error);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [includeTasks]);

  // Transform events for FullCalendar
  const fullCalendarEvents: FullCalendarEvent[] = events.map((ev) => ({
    id: ev.id,
    title: ev.title,
    start: ev.start,
    end: ev.end,
    color: ev.resource
      ? tokens.colorPaletteBlueBackground2
      : tokens.colorPaletteGreenBackground2,
    allDay: ev.isAllDay,
    extendedProps: {
      resource: ev.resource,
      organizer: ev.organizer,
      eventType: ev.eventType,
    },
  }));

  // Custom event content for FullCalendar
  function renderEventContent(eventInfo: EventContentArg): JSX.Element {
    const task = eventInfo.event.extendedProps.resource as TaskItem | undefined;
    return (
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 4,
          width: "100%",
          overflow: "hidden",
        }}
      >
        {task?.AssignedTo && (
          <Persona
            size="extra-small"
            avatar={{
              image: {
                src: `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(
                  task.AssignedTo.EMail
                )}`,
                alt: task.AssignedTo.Title,
              },
            }}
          />
        )}
        <span
          style={{
            whiteSpace: "nowrap",
            overflow: "hidden",
            textOverflow: "ellipsis",
          }}
        >
          {eventInfo.event.title}
        </span>
      </div>
    );
  }

  const handleEventChange = async (
    changeInfo: EventChangeArg
  ): Promise<void> => {
    const { event } = changeInfo;
    const eventType = event.extendedProps.eventType;
    if (eventType === "Task") {
      const task = event.extendedProps.resource as TaskItem | undefined;
      console.log("event", event);
      if (task) {
        try {
          // Update the task in SharePoint
          await tasksService.updateTask(props.listName, {
            ...task,
            StartDate: event.start?.toISOString() || new Date().toISOString(),
            DueDate: event.end?.toISOString() || new Date().toISOString(),
            Priority: task.Priority || "Medium",
            Status: task.Status || "Not Started",
            ProjectID: task.ProjectID || "",
          });

          // Force a complete reload of events
          // await loadEvents();
          // Update the events state with the new task dates
          setEvents((prevEvents) => {
            return prevEvents.map((evt) => {
              if (evt.id === event.id) {
                return {
                  ...evt,
                  start: event.start || evt.start,
                  end: event.end || evt.end,
                };
              }
              return evt;
            });
          });

          // Force calendar to re-render by updating the key
          const calendarKey = Date.now();
          setCalendarKey(calendarKey);
        } catch (err) {
          setError(err.message);
          // Revert the event if there was an error
          changeInfo.revert();
        }
      }
    } else if (eventType === "Calender") {
      // This is a calendar event (Outlook/Graph)
      // TODO: Add your update logic for Graph events here, e.g. PATCH to /me/events/{id}
      const client =
        graphClient ||
        (await props.context.msGraphClientFactory.getClient("3"));
      if (!graphClient) setGraphClient(client);
      if (!event.start || !event.end) {
        changeInfo.revert();
        return;
      }
      await client.api(`/me/events/${event.id}`).patch({
        start: { dateTime: event.start.toISOString(), timeZone: "UTC" },
        end: { dateTime: event.end.toISOString(), timeZone: "UTC" },
      });
      // Then update state as needed
      // For now, just update the state locally:
      setEvents((prevEvents) =>
        prevEvents.map((evt) =>
          evt.id === event.id
            ? {
                ...evt,
                start: event.start || evt.start,
                end: event.end || evt.end,
              }
            : evt
        )
      );
      setCalendarKey(Date.now());
    }
  };

  const handleEventClick = (clickInfo: EventClickArg): void => {
    console.log("clickInfo", clickInfo);
    setSelectedEvent(clickInfo.event);
    setIsModalOpen(true);
  };

  if (isLoading) {
    return <Spinner size="large" label="Loading calendar..." />;
  }

  if (error) {
    return (
      <Text size={400} style={{ color: tokens.colorPaletteRedForeground1 }}>
        {error}
      </Text>
    );
  }

  function eventCard(event: CalendarEventForDisplay): JSX.Element {
    return (
      <Tooltip
        content={`${event.subject}\nOrganizer: ${event.organizer.emailAddress.name}`}
        relationship="label"
        key={event.id}
      >
        <div className={styles.eventCard}>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              width: "100%",
            }}
          >
            {event.ProjectID ? (
              <Text size={300} weight="semibold">
                Project: {event.ProjectID}
              </Text>
            ) : (
              <span />
            )}
            {event.eventType && <Text size={200}>Type: {event.eventType}</Text>}
          </div>
          <div>
            <Text size={300} weight="bold">
              {event.title || event.subject}
            </Text>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Text size={200}>Start: {event.start.toDateString()}</Text>
            <Text size={200}>End: {event.end.toDateString()}</Text>
          </div>
          <div>
            <Text size={200}>
              {event.isAllDay
                ? "All Day"
                : `Time: ${event.start.toLocaleTimeString()} - ${event.end.toLocaleTimeString()}`}
            </Text>
          </div>
        </div>
      </Tooltip>
    );
  }

  return (
    <>
      <div className={stylesFluent.main}>
        <div className={stylesFluent.calendarContainer}>
          <div className={stylesFluent.controls}>
            <Switch
              label="Include SharePoint Tasks"
              checked={includeTasks}
              onChange={(_, data) => setIncludeTasks(data.checked)}
            />
            <Switch
              label="Show Weekends"
              checked={showWeekends}
              onChange={(_, data) => setShowWeekends(data.checked)}
            />
            <Switch
              label="Show Business Hours (8am-6pm)"
              checked={showBusinessHours}
              onChange={(_, data) => setShowBusinessHours(data.checked)}
            />
          </div>
          <FullCalendar
            key={calendarKey}
            plugins={[
              dayGridPlugin,
              timeGridPlugin,
              interactionPlugin,
              listPlugin,
              multiMonthPlugin,
            ]}
            initialView="dayGridMonth"
            weekends={showWeekends}
            events={fullCalendarEvents}
            height={600}
            eventContent={renderEventContent}
            editable={true}
            droppable={true}
            eventDrop={handleEventChange}
            eventResize={handleEventChange}
            navLinks={true}
            businessHours={{
              daysOfWeek: [1, 2, 3, 4, 5],
              startTime: "08:00",
              endTime: "18:00",
            }}
            slotMinTime={showBusinessHours ? "08:00" : "00:00"}
            slotMaxTime={showBusinessHours ? "18:00" : "23:59"}
            eventClick={handleEventClick}
            views={{
              dayGridMonth: {},
              dayGridWeek: {},
              dayGridDay: {},
              timeGridWeek: {},
              timeGridDay: {},
              listYear: {},
              listMonth: {},
              listWeek: {},
              listDay: {},
              multiMonthYear: {},
              multiMonth: {},
            }}
            headerToolbar={{
              left: "prev,next today",
              center: "title",
              right: "dayGridMonth,timeGridWeek,timeGridDay,listWeek",
            }}
          />
        </div>
        <div className={stylesFluent.rightPanel}>
          <Text size={400}>Calendar Events</Text>

          {/* Upcoming Events */}
          <Text size={300}>Upcoming Events</Text>
          {events
            .filter((event) => event.start > new Date())
            .sort((a, b) => a.start.getTime() - b.start.getTime())
            .map((event) => eventCard(event))}
          {/* Past Events */}
          <Text size={300}>Past Events</Text>
          {events
            .filter((event) => event.start <= new Date())
            .sort((a, b) => b.start.getTime() - a.start.getTime()) // Reverse chronological
            .map((event) => eventCard(event))}
        </div>
      </div>
      {/* Move Dialog here, outside the main layout */}
      {selectedEvent && (
        <TaskAddEditDialog
          context={props.context}
          open={isModalOpen}
          onClose={() => setIsModalOpen(false)}
          onSave={() => {}} // <-- no-op function TO DO: Add save logic
          mode="view"
          task={selectedEvent?.extendedProps?.resource as ITask}
        />
      )}
    </>
  );
};
