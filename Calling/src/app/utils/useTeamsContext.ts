/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import * as microsoftTeams from '@microsoft/teams-js';
import { app } from '@microsoft/teams-js';
import { useEffect, useState } from 'react';
import { inTeams } from '../utils/inTeams';

export const useTeamsContext = (initialized: boolean) => {
  const [ctx, setCtx] = useState<app.Context | null>(null);

  useEffect(() => {
    if (!ctx?.user?.id) {
      // Add inTeams=true to URL params to get real Teams context
      if (inTeams() && initialized) {
        console.log('useTeamsContext: Attempting to get Teams context');
        // Get Context from the Microsoft Teams SDK
        microsoftTeams.app
          .getContext()
          .then((context) => {
            console.log(`useTeamsContext: received context: ${JSON.stringify(context)}`);
            setCtx(context);
          })
          .catch((error) => console.error(error));
      }
    }
  }, [ctx?.user?.id, initialized]);

  return ctx;
};
