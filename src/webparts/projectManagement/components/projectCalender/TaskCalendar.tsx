import React, { useState } from 'react';
import {
  Calendar as BigCalendar,
  momentLocalizer,
  Event as CalendarEvent,
  Views,
} from 'react-big-calendar';
import {
  makeStyles,
  tokens,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Persona,
} from '@fluentui/react-components';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import { parseISO } from 'date-fns';
import moment from 'moment';
import { TaskItem } from './ProgrammeTab';

const localizer = momentLocalizer(moment);

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    height: '100%',
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  calendar: {
    height: '600px',
    backgroundColor: '#fff',
    padding: '20px',
    borderRadius: '12px',
  },
  dialogContent: {
    display: 'grid',
    gap: tokens.spacingVerticalS,
  },
  eventWrapper: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    width: '100%',
    overflow: 'hidden',
  },
  eventTitle: {
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  }
});

interface Props {
  tasks: TaskItem[];
  onTaskClick?: (task: TaskItem) => void;
}

const EventComponent = ({ event }: { event: CalendarEvent }): JSX.Element => {
  const styles = useStyles();
  const task = event.resource as TaskItem;
  
  return (
    <div className={styles.eventWrapper}>
      {task.AssignedTo && (
        <Persona
          size="extra-small"
          avatar={{
            image: {
              src: `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(task.AssignedTo.EMail)}`,
              alt: task.AssignedTo.Title,
            }
          }}
        />
      )}
      <span className={styles.eventTitle}>{event.title}</span>
    </div>
  );
};

const TaskCalendar: React.FC<Props> = ({ tasks, onTaskClick }) => {
  const styles = useStyles();
  const [selectedEvent, setSelectedEvent] = useState<CalendarEvent | null>(null);
  const [showEventDialog, setShowEventDialog] = useState(false);

  const calendarEvents: CalendarEvent[] = tasks.map((task) => ({
    id: task.Id,
    title: task.Title,
    start: task.StartDate ? parseISO(task.StartDate) : new Date(),
    end: task.DueDate ? parseISO(task.DueDate) : new Date(),
    resource: task,
  }));

  const handleSelectEvent = (event: CalendarEvent): void => {
    setSelectedEvent(event);
    setShowEventDialog(false);
    if (onTaskClick && event.resource) {
      onTaskClick(event.resource);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <h2>Project Calendar</h2>
      </div>

      <BigCalendar
        localizer={localizer}
        events={calendarEvents}
        startAccessor="start"
        endAccessor="end"
        style={{ height: '100%' }}
        views={[Views.MONTH, Views.WEEK, Views.DAY]}
        defaultView={Views.MONTH}
        onSelectEvent={handleSelectEvent}
        className={styles.calendar}
        components={{
          event: EventComponent
        }}
        popup
      />

      <Dialog open={showEventDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              Task Details
            </DialogTitle>
            <DialogContent>
              <div className={styles.dialogContent}>
                {selectedEvent && (
                  <>
                    <h3>{selectedEvent.title}</h3>
                    <p>{selectedEvent.resource.Description}</p>
                    <p>
                      {moment(selectedEvent.start).format('dddd, MMM Do YYYY h:mm A')} —{' '}
                      {moment(selectedEvent.end).format('h:mm A')}
                    </p>
                  </>
                )}
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => setShowEventDialog(false)}
              >
                Close
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default TaskCalendar;
