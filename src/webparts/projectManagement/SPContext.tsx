import * as React from 'react';
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export const SPContext = React.createContext<SPFI>(null!);

export const SPProvider: React.FC<{ context: WebPartContext, children: React.ReactNode }> = ({ context, children }) => {
  const sp = React.useMemo(() => spfi().using(SPFx(context)), [context]);
  return <SPContext.Provider value={sp}>{children}</SPContext.Provider>;
};
