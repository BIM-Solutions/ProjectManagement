import * as React from "react";
import { useState, useEffect } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ITask } from "../services/TasksService";
import TasksService from "../services/TasksService";
import { TaskCard } from "./TaskCard";
import { TaskDetailsDialog } from "./TaskDetailsDialog";
import { Spinner } from "@fluentui/react-components";
import {
  DragDropContext,
  Droppable,
  Draggable,
  DropResult,
  DroppableProvided,
  DroppableStateSnapshot,
  DraggableProvided,
  DraggableStateSnapshot,
} from "react-beautiful-dnd";

export interface TaskBoardProps {
  listName: string;
  userDisplayName: string;
  context: WebPartContext;
}

const statusColumns = [
  { key: "Not Started", title: "Not Started" },
  { key: "In Progress", title: "In Progress" },
  { key: "Completed", title: "Completed" },
];

export const TaskBoard: React.FC<TaskBoardProps> = ({
  listName,
  userDisplayName,
  context,
}) => {
  const [tasks, setTasks] = useState<ITask[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedTask, setSelectedTask] = useState<ITask | null>(null);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const tasksService = TasksService.getInstance(context);

  const loadTasks = async (): Promise<void> => {
    setIsLoading(true);
    try {
      const items = await tasksService.getTasks(listName, userDisplayName);
      setTasks(items);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadTasks().catch((err) => console.log(err));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [listName, userDisplayName]);

  const onDragEnd = async (result: DropResult): Promise<void> => {
    if (!result.destination) return;
    const { source, destination, draggableId } = result;
    if (source.droppableId === destination.droppableId) return;
    const taskId = parseInt(draggableId, 10);
    const task = tasks.find((t) => t.Id === taskId);
    if (!task) return;
    try {
      await tasksService.updateTaskStatus(
        listName,
        taskId,
        destination.droppableId,
        task.AssignedTo?.Title
      );
      await loadTasks();
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  const groupedTasks = statusColumns.reduce((acc, col) => {
    acc[col.key] = tasks.filter((t) => t.Status === col.key);
    return acc;
  }, {} as Record<string, ITask[]>);

  if (isLoading) return <Spinner label="Loading tasks..." />;
  if (error) return <div>{error}</div>;

  return (
    <DragDropContext onDragEnd={onDragEnd}>
      <div style={{ display: "flex", gap: 24, width: "100%" }}>
        {statusColumns.map((col) => (
          <Droppable droppableId={col.key} key={col.key}>
            {(
              provided: DroppableProvided,
              snapshot: DroppableStateSnapshot
            ) => (
              <div
                ref={provided.innerRef}
                {...provided.droppableProps}
                style={{
                  flex: 1,
                  background: snapshot.isDraggingOver ? "#f3f2f1" : "#faf9f8",
                  borderRadius: 12,
                  minHeight: 500,
                  padding: 16,
                  boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
                  transition: "background 0.2s",
                }}
              >
                <h3 style={{ marginBottom: 16 }}>{col.title}</h3>
                {groupedTasks[col.key].map((task, idx) => (
                  <Draggable
                    draggableId={task.Id.toString()}
                    index={idx}
                    key={task.Id}
                  >
                    {(
                      provided: DraggableProvided,
                      snapshot: DraggableStateSnapshot
                    ) => (
                      <div
                        ref={provided.innerRef}
                        {...provided.draggableProps}
                        {...provided.dragHandleProps}
                        style={{
                          marginBottom: 16,
                          ...provided.draggableProps.style,
                        }}
                      >
                        <TaskCard
                          task={task}
                          onClick={() => {
                            setSelectedTask(task);
                            setIsDialogOpen(true);
                          }}
                        />
                      </div>
                    )}
                  </Draggable>
                ))}
                {provided.placeholder}
              </div>
            )}
          </Droppable>
        ))}
      </div>
      {selectedTask && (
        <TaskDetailsDialog
          task={selectedTask}
          open={isDialogOpen}
          onClose={() => setIsDialogOpen(false)}
        />
      )}
    </DragDropContext>
  );
};
