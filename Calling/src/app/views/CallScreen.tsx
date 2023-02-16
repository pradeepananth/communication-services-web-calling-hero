// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CommunicationUserIdentifier } from '@azure/communication-common';
import {
  CallAdapterLocator,
  CallAdapter,
  CallAdapterState,
  CallComposite,
  toFlatCommunicationIdentifier,
  useAzureCommunicationCallAdapter
} from '@azure/communication-react';

import { Spinner, Stack } from '@fluentui/react';
import React, { memo, useCallback, useEffect, useMemo, useRef } from 'react';
import { useSwitchableFluentTheme } from '../theming/SwitchableFluentThemeProvider';
import { createAutoRefreshingCredential } from '../utils/credential';
import { WEB_APP_TITLE } from '../utils/AppUtils';
import { useIsMobile } from '../utils/useIsMobile';
import { NoteContainer } from '../NoteContainer';
import { TeamsMeetingLinkLocator } from '@azure/communication-calling';

export interface CallScreenProps {
  token: string;
  userId: CommunicationUserIdentifier;
  callLocator: CallAdapterLocator;
  displayName: string;

  onCallEnded: () => void;
}

export const CallScreen = memo((props: CallScreenProps): JSX.Element => {
  const { token, userId, callLocator, displayName, onCallEnded } = props;
  const callIdRef = useRef<string>();
  const [callStatus, setCallStatus] = React.useState<string>('');
  const { currentTheme, currentRtl } = useSwitchableFluentTheme();
  const isMobileSession = useIsMobile();
  const afterCreate = useCallback(
    async (adapter: CallAdapter): Promise<CallAdapter> => {
      adapter.on('callEnded', () => {
        onCallEnded();
      });
      adapter.on('error', (e) => {
        // Error is already acted upon by the Call composite, but the surrounding application could
        // add top-level error handling logic here (e.g. reporting telemetry).
        console.log('Adapter error event:', e);
      });
      adapter.onStateChange((state: CallAdapterState) => {
        const pageTitle = convertPageStateToString(state);
        setCallStatus(pageTitle);
        document.title = `${pageTitle} - ${WEB_APP_TITLE}`;

        if (state?.call?.id && callIdRef.current !== state?.call?.id) {
          callIdRef.current = state?.call?.id;
          console.log(`Call Id: ${callIdRef.current}`);
        }
      });
      return adapter;
    },
    [callIdRef, onCallEnded]
  );

  const credential = useMemo(
    () => createAutoRefreshingCredential(toFlatCommunicationIdentifier(userId), token),
    [token, userId]
  );

  const adapter = useAzureCommunicationCallAdapter(
    {
      userId,
      displayName,
      credential,
      locator: callLocator
    },

    afterCreate
  );

  // Dispose of the adapter in the window's before unload event.
  // This ensures the service knows the user intentionally left the call if the user
  // closed the browser tab during an active call.
  useEffect(() => {
    const disposeAdapter = (): void => adapter?.dispose();
    window.addEventListener('beforeunload', disposeAdapter);
    return () => window.removeEventListener('beforeunload', disposeAdapter);
  }, [adapter]);

  if (!adapter) {
    return <Spinner label={'Creating adapter'} ariaLive="assertive" labelPosition="top" />;
  }

  const callInvitationUrl: string | undefined = window.location.href;

  return (
    <Stack verticalFill>
      {callStatus === 'call' && (
        <Stack styles={{ root: { height: '58vh', overflow: 'auto' } }}>
          <NoteContainer
            acsLiveShareHostOptions={{
              callAdapter: adapter,
              teamsMeetingJoinUrl: (callLocator as TeamsMeetingLinkLocator).meetingLink,
              acsTokenProvider: () => token
            }}
          />
        </Stack>
      )}
      <Stack styles={{ root: { height: '18vh' } }}>
        <CallComposite
          adapter={adapter}
          fluentTheme={currentTheme.theme}
          rtl={currentRtl}
          callInvitationUrl={callInvitationUrl}
          formFactor={'desktop'}
          options={{
            errorBar: true,
            callControls: {
              cameraButton: true,
              microphoneButton: true,
              screenShareButton: true,
              endCallButton: true,
              participantsButton: false
            }
          }}
        />
      </Stack>
    </Stack>
  );
});

const convertPageStateToString = (state: CallAdapterState): string => {
  switch (state.page) {
    case 'accessDeniedTeamsMeeting':
      return 'error';
    case 'leftCall':
      return 'end call';
    case 'removedFromCall':
      return 'end call';
    default:
      return `${state.page}`;
  }
};
