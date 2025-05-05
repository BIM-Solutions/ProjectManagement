import React, { useState} from 'react';
import {
  Calendar as BigCalendar,
  momentLocalizer,
  Event as CalendarEvent,
} from 'react-big-calendar';
import { Modal, Text } from '@fluentui/react';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import { parseISO } from 'date-fns';
import moment from 'moment';
import { TaskItem } from './ProgrammeTab'; // Adjust the import path as necessary


const localizer = momentLocalizer(moment);



interface Props {
  tasks: TaskItem[];
  onTaskClick?: (task: TaskItem) => void;
}

const TaskCalendar: React.FC<Props> = ({ tasks, onTaskClick }) => {
  const [selectedEvent, setSelectedEvent] = useState<CalendarEvent | null>(null);
  const [modalVisible, setModalVisible] = useState(false);

  const calendarEvents: CalendarEvent[] = tasks.map((task) => ({
    id: task.Id,
    title: task.Title,
    start: task.StartDate ? parseISO(task.StartDate) : new Date(),
    end: task.DueDate ? parseISO(task.DueDate) : new Date(),
    resource: task,
  }));

  const handleSelectEvent = (event: CalendarEvent): void => {
    setSelectedEvent(event);
    setModalVisible(true);
  };

  return (
    <div style={{ height: '700px', backgroundColor: '#fff', padding: 20, borderRadius: 12 }}>
      <BigCalendar
        localizer={localizer}
        events={calendarEvents}
        startAccessor="start"
        endAccessor="end"
        style={{ height: '100%', borderRadius: 12 }}
        views={['month']}
        onSelectEvent={handleSelectEvent}
        popup
      />

      <Modal
        isOpen={modalVisible}
        onDismiss={() => setModalVisible(false)}
        isBlocking={false}
      >
        {selectedEvent && (
          <div style={{ padding: 20 }}>
            <Text variant="xLarge">{selectedEvent.title}</Text>
            <Text>{selectedEvent.resource.Description}</Text>
            <Text>
              {moment(selectedEvent.start).format('dddd, MMM Do YYYY h:mm A')} —{' '}
              {moment(selectedEvent.end).format('h:mm A')}
            </Text>
          </div>
        )}
      </Modal>
    </div>
  );
};

export default TaskCalendar;
