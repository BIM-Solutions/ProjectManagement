import * as React from "react";
import { useState, useEffect } from "react";
import {
  Stack,
  Text,
  Toggle,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TooltipHost,
} from "@fluentui/react";
import { Calendar, DateRangeType } from "@fluentui/react-calendar-compat";
import styles from "./CalendarView.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";

interface ICalendarEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  organizer: { emailAddress: { name: string; address: string } };
}

interface ICalendarViewProps {
  showTeamCalendar: boolean;
  groupId: string;
  context: WebPartContext;
}

export const CalendarView: React.FC<ICalendarViewProps> = (props) => {
  const [events, setEvents] = useState<ICalendarEvent[]>([]);
  const [teamEvents, setTeamEvents] = useState<ICalendarEvent[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const [viewType, setViewType] = useState<DateRangeType>(DateRangeType.Week);
  const [showTeam, setShowTeam] = useState(props.showTeamCalendar);
  const [graphClient, setGraphClient] = useState<MSGraphClientV3 | null>(null);

  useEffect(() => {
    props.context.msGraphClientFactory
      .getClient("3")
      .then((client) => setGraphClient(client))
      .catch(console.error);
  }, [props.context]);

  const viewOptions: IDropdownOption[] = [
    { key: DateRangeType.Day, text: "Day" },
    { key: DateRangeType.Week, text: "Week" },
    { key: DateRangeType.Month, text: "Month" },
  ];

  const loadEvents = async (): Promise<void> => {
    try {
      setIsLoading(true);

      if (!graphClient) {
        const client = await props.context.msGraphClientFactory.getClient("3");
        setGraphClient(client);
        if (showTeam) {
          const teamResponse = await client
            .api(`/groups/${props.groupId}/calendar/events`)
            .get();
          setTeamEvents(teamResponse.value);
        } else {
          const personalResponse = await client
            .api("/me/calendar/events")
            .get();
          setEvents(personalResponse.value);
        }
      } else {
        if (showTeam) {
          const teamResponse = await graphClient
            .api(`/groups/${props.groupId}/calendar/events`)
            .get();
          setTeamEvents(teamResponse.value);
        } else {
          const personalResponse = await graphClient
            .api("/me/calendar/events")
            .get();
          setEvents(personalResponse.value);
        }
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  };
  useEffect(() => {
    loadEvents().catch(console.error);
  }, [showTeam]);

  const renderEvent = (event: ICalendarEvent): JSX.Element => (
    <TooltipHost
      content={`${event.subject}\nOrganizer: ${event.organizer.emailAddress.name}`}
      key={event.id}
    >
      <div className={styles.eventCard}>
        <Text variant="small">{event.subject}</Text>
        <Text variant="small">
          {new Date(event.start.dateTime).toLocaleTimeString()} -
          {new Date(event.end.dateTime).toLocaleTimeString()}
        </Text>
      </div>
    </TooltipHost>
  );

  if (isLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading calendar..." />;
  }

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 15 }}>
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <Toggle
          label="Show Team Calendar"
          checked={showTeam}
          onText="Team"
          offText="Personal"
          onChange={(_, checked) => setShowTeam(checked ?? false)}
        />
        <Dropdown
          label="View"
          options={viewOptions}
          selectedKey={viewType}
          onChange={(_, item) => item && setViewType(item.key as DateRangeType)}
        />
      </Stack>

      <Calendar
        value={selectedDate}
        onSelectDate={setSelectedDate}
        dateRangeType={viewType}
        showGoToToday
        highlightSelectedMonth
        showWeekNumbers
        firstDayOfWeek={1}
      />

      <Stack tokens={{ childrenGap: 10 }}>
        <Text variant="large">Events</Text>
        {showTeam ? teamEvents.map(renderEvent) : events.map(renderEvent)}
      </Stack>
    </Stack>
  );
};
