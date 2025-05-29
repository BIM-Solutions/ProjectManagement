import { useEffect, useState } from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const useGraph = (
  context?: WebPartContext
): { graphClient: MSGraphClientV3 | undefined } => {
  const [graphClient, setGraphClient] = useState<MSGraphClientV3 | undefined>(
    undefined
  );

  useEffect(() => {
    if (context) {
      context.msGraphClientFactory
        .getClient("3")
        .then((client) => {
          setGraphClient(client);
        })
        .catch(console.error);
    }
  }, [context]);

  return { graphClient };
};
