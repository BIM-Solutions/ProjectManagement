import * as React from "react";
import { useState, useEffect } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ITask } from "../../services/TasksService";
import TasksService from "../../services/TasksService";
import { TaskCard } from "./TaskCard";
// import { TaskDetailsDialog } from "./TaskDetailsDialog";
import {
  Spinner,
  Button,
  // Dialog,
  // DialogSurface,
  // DialogTitle,
  // DialogBody,
  // DialogActions,
  // Input,
  // Dropdown,
  // Option,
} from "@fluentui/react-components";
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
import { TaskAddEditDialog } from "../../../common/components/Event/TaskAddEditDialog";

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
  const [isAddDialogOpen, setIsAddDialogOpen] = useState(false);
  const tasksService = TasksService.getInstance(context);

  const loadTasks = async (): Promise<void> => {
    setIsLoading(true);
    try {
      const items = await tasksService.getTasks(listName, userDisplayName);
      console.log(items);
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
    // Update tasks state with new status
    setTasks(
      tasks.map((t) => {
        if (t.Id === taskId) {
          return { ...t, Status: destination.droppableId };
        }
        return t;
      })
    );
    try {
      await tasksService.updateTaskStatus(
        listName,
        taskId,
        destination.droppableId,
        task.AssignedTo?.Title
      );
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  const groupedTasks = statusColumns.reduce((acc, col) => {
    acc[col.key] = tasks.filter((t) => t.Status === col.key);
    return acc;
  }, {} as Record<string, ITask[]>);

  const handleAddTask = async (task: Partial<ITask>): Promise<void> => {
    await tasksService.addTask(listName, task, userDisplayName);
    setIsAddDialogOpen(false);
    await loadTasks();
  };

  if (isLoading) return <Spinner label="Loading tasks..." />;
  if (error) return <div>{error}</div>;

  return (
    <>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 16,
        }}
      >
        <h2 style={{ margin: 0 }}>Tasks Board</h2>
        <Button appearance="primary" onClick={() => setIsAddDialogOpen(true)}>
          + New Task
        </Button>
      </div>
      <DragDropContext onDragEnd={onDragEnd}>
        <div
          style={{
            display: "flex",
            justifyContent: "flex-start",

            gap: 24,
            width: "100%",
          }}
        >
          {statusColumns.map((col) => (
            <div
              key={col.key}
              style={{
                flex: 1,
                maxWidth: 400,
              }}
            >
              <h3 style={{ marginBottom: 16, marginLeft: 16, maxWidth: 400 }}>
                {col.title}
              </h3>
              <Droppable droppableId={col.key}>
                {(
                  provided: DroppableProvided,
                  snapshot: DroppableStateSnapshot
                ) => (
                  <div
                    ref={provided.innerRef}
                    {...provided.droppableProps}
                    style={{
                      flex: 1,
                      background: snapshot.isDraggingOver
                        ? "#f3f2f1"
                        : "#faf9f8",
                      borderRadius: 12,
                      minHeight: 500,
                      maxHeight: 600,
                      maxWidth: 400,
                      overflowY: "auto",
                      padding: 16,
                      boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
                      transition: "background 0.2s",
                    }}
                  >
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
            </div>
          ))}
        </div>
      </DragDropContext>
      {selectedTask && (
        <TaskAddEditDialog
          context={context}
          open={isDialogOpen}
          onClose={() => setIsDialogOpen(false)}
          onSave={handleAddTask}
          mode="view"
          task={selectedTask}
        />
      )}
      <TaskAddEditDialog
        context={context}
        open={isAddDialogOpen}
        onClose={() => setIsAddDialogOpen(false)}
        onSave={handleAddTask}
        mode="add"
      />
    </>
  );
};
