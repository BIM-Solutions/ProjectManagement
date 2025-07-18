import { useEffect, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export const usePnP = (context?: WebPartContext): { sp: SPFI | undefined } => {
  const [sp, setSp] = useState<SPFI | undefined>(undefined);

  useEffect(() => {
    if (context) {
      const spInstance = spfi().using(SPFx(context));
      setSp(spInstance);
    }
  }, [context]);

  return { sp };
};
