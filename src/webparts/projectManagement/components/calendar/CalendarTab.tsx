import * as React from 'react';
import { useState, useEffect } from 'react';
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
} from '@fluentui/react-components';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IProject } from '../../services/ProjectService';
import { Calendar, Views, momentLocalizer } from 'react-big-calendar';
import moment from 'moment';
import 'react-big-calendar/lib/css/react-big-calendar.css';

const localizer = momentLocalizer(moment);

export interface ICalendarTabProps {
  context: WebPartContext;
  project: IProject | undefined;
}

interface ICalendarEvent {
  id: number;
  title: string;
  start: Date;
  end: Date;
  description: string;
  type: string;
  status: string;
}

const eventTypes = [
  'Meeting',
  'Deadline',
  'Review',
  'Site Visit',
  'CofW',
  'Other',
];

const eventStatuses = [
  'Scheduled',
  'Completed',
  'Cancelled',
  'Pending',
];

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
  },
  form: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalM,
  },
  fullWidth: {
    gridColumn: '1 / -1',
  },
});

const CalendarTab: React.FC<ICalendarTabProps> = ({ context, project }) => {
  const styles = useStyles();
  const [events, setEvents] = useState<ICalendarEvent[]>([]);
  const [showEventDialog, setShowEventDialog] = useState(false);
  const [selectedEvent, setSelectedEvent] = useState<ICalendarEvent | null>(null);
  const [formData, setFormData] = useState<Partial<ICalendarEvent>>({});
  // const [selectedDate, setSelectedDate] = useState<Date | null>(null);

  const loadEvents = async (): Promise<void> => {
    try {
      // This would be replaced with actual API call to load events
      const mockEvents: ICalendarEvent[] = [
        {
          id: 1,
          title: 'Project Review',
          start: new Date(),
          end: new Date(Date.now() + 2 * 60 * 60 * 1000), // 2 hours later
          description: 'Monthly project review meeting',
          type: 'Meeting',
          status: 'Scheduled',
        },
      ];
      setEvents(mockEvents);
    } catch (error) {
      console.error('Error loading events:', error);
    }
  };

  useEffect(() => {
    if (project) {
      loadEvents().catch(console.error);
    }
  }, [project]);

  const handleInputChange = (
    field: keyof ICalendarEvent,
    value: string | Date
  ): void => {
    setFormData(prev => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleSave = async (): Promise<void>   => {
    try {
      if (selectedEvent) {
        // Update existing event
        const updatedEvents = events.map(event =>
          event.id === selectedEvent.id ? { ...event, ...formData } : event
        );
        setEvents(updatedEvents);
      } else {
        // Create new event
        const newEvent = {
          id: Math.max(...events.map(e => e.id), 0) + 1,
          ...formData,
        } as ICalendarEvent;
        setEvents([...events, newEvent]);
      }
      setShowEventDialog(false);
      setSelectedEvent(null);
      setFormData({});
    } catch (error) {
      console.error('Error saving event:', error);
    }
  };

  const handleSelectSlot = ({ start, end }: { start: Date; end: Date }): void => {
    setFormData({
      start,
      end,
    });
    setShowEventDialog(true);
  };

  const handleSelectEvent = (event: ICalendarEvent): void => {
    setSelectedEvent(event);
    setFormData(event);
    setShowEventDialog(true);
  };

  return (
    <div className={styles.container}>
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

      <Dialog open={showEventDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              {selectedEvent ? 'Edit Event' : 'New Event'}
            </DialogTitle>
            <DialogContent>
              <div className={styles.form}>
                <Field label="Title" required>
                  <Input
                    value={formData.title || ''}
                    onChange={(_, data) => handleInputChange('title', data.value)}
                  />
                </Field>

                <Field label="Type">
                  <Dropdown
                    value={formData.type || ''}
                    onOptionSelect={(_, data) =>
                      handleInputChange('type', data.optionValue || '')
                    }
                  >
                    {eventTypes.map((type) => (
                      <Option key={type} value={type}>
                        {type}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Status">
                  <Dropdown
                    value={formData.status || ''}
                    onOptionSelect={(_, data) =>
                      handleInputChange('status', data.optionValue || '')
                    }
                  >
                    {eventStatuses.map((status) => (
                      <Option key={status} value={status}>
                        {status}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Start Date">
                  <DatePicker
                    value={formData.start ? new Date(formData.start) : undefined}
                    onSelectDate={(date: Date | null | undefined) => {
                      if (date) {
                        handleInputChange('start', date);
                      }
                    }}
                  />
                </Field>

                <Field label="End Date">
                  <DatePicker
                    value={formData.end ? new Date(formData.end) : undefined}
                    onSelectDate={(date: Date | null | undefined) => {
                      if (date) {
                        handleInputChange('end', date);
                      }
                    }}
                  />
                </Field>

                <Field label="Description" className={styles.fullWidth}>
                  <Textarea
                    value={formData.description || ''}
                    onChange={(_, data) =>
                      handleInputChange('description', data.value)
                    }
                  />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => setShowEventDialog(false)}
              >
                Cancel
              </Button>
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