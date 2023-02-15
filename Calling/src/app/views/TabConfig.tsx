/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { Stack } from '@fluentui/react';
import * as microsoftTeams from '@microsoft/teams-js';
import { containerStyles, headerStyles } from '../styles/CallScreen.styles';
import React, { FC, useEffect } from 'react';

const TabConfig: FC = () => {
  useEffect(() => {
    microsoftTeams.pages.config.registerOnSaveHandler(function (saveEvent) {
      microsoftTeams.pages.config.setConfig({
        suggestedDisplayName: 'ACS sample app',
        contentUrl: `${window.location.origin}?inTeams=true`
      });
      saveEvent.notifySuccess();
    });

    microsoftTeams.pages.config.setValidityState(true);
  }, []);

  return (
    <Stack styles={containerStyles}>
      <Stack styles={headerStyles}>Welcome to ACS Live Share Demo!</Stack>
      <Stack as="p">Press the save button to continue.</Stack>
    </Stack>
  );
};

export default TabConfig;
