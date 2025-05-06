const PROJECT_UPDATED_EVENT = 'projectUpdated';

export const EventService = {
  publishProjectUpdated: () => {
    window.dispatchEvent(new CustomEvent(PROJECT_UPDATED_EVENT));
  },
  subscribeToProjectUpdates: (callback: () => void) => {
    window.addEventListener(PROJECT_UPDATED_EVENT, callback);
    return () => window.removeEventListener(PROJECT_UPDATED_EVENT, callback);
  }
};
