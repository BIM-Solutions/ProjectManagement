const PROJECT_UPDATED_EVENT = 'projectUpdated';

type EventCallback = () => void;

class EventService {
  private static instance: EventService;
  private documentUploadListeners: EventCallback[] = [];

  private constructor() {}

  static getInstance(): EventService {
    if (!EventService.instance) {
      EventService.instance = new EventService();
    }
    return EventService.instance;
  }

  subscribeToDocumentUpload(callback: EventCallback): void {
    this.documentUploadListeners.push(callback);
  }

  unsubscribeFromDocumentUpload(callback: EventCallback): void {
    this.documentUploadListeners = this.documentUploadListeners.filter(cb => cb !== callback);
  }

  notifyDocumentUpload(): void {
    this.documentUploadListeners.forEach(callback => callback());
  }

  publishProjectUpdated: () => void = () => {
    window.dispatchEvent(new CustomEvent(PROJECT_UPDATED_EVENT));
  };

  subscribeToProjectUpdates: (callback: () => void) => () => void = (callback) => {
    window.addEventListener(PROJECT_UPDATED_EVENT, callback);
    return () => window.removeEventListener(PROJECT_UPDATED_EVENT, callback);
  };
}

export const eventService = EventService.getInstance();
