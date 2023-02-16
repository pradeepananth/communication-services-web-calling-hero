/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { PrimaryButton, Stack, Text } from '@fluentui/react';
import * as microsoftTeams from '@microsoft/teams-js';
import { buttonWithIconStyles } from '../styles/Footer.styles';
import { buttonStyle } from '../styles/StartCallButton.styles';
import React, { useCallback } from 'react';

export const SidePanel = () => {
  const shareToStage = useCallback(() => {
    microsoftTeams.meeting.shareAppContentToStage((error) => {
      if (error) {
        console.error(error);
      }
    }, `${window.location.origin}/?inTeams=true#1d3e40ee-98d1-45df-a44d-d28d63f07c75`);
  }, []);

  return (
    <Stack horizontal wrap horizontalAlign="center" verticalAlign="center">
      <Stack>
        <Text role={'heading'} aria-level={1}>
          {'Please share the app to the stage'}
        </Text>
        <Stack horizontal>
          <PrimaryButton
            className={buttonStyle}
            styles={buttonWithIconStyles}
            text={'Share to stage'}
            onClick={shareToStage}
          />
        </Stack>
      </Stack>
    </Stack>
  );
};
