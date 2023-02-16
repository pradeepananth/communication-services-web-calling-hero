// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CommunicationUserIdentifier } from '@azure/communication-common';

import { setLogLevel } from '@azure/logger';
import { initializeIcons, Spinner } from '@fluentui/react';
import { CallAdapterLocator } from '@azure/communication-react';
import React, { useEffect, useState } from 'react';
import {
  createGroupId,
  fetchTokenResponse,
  getGroupIdFromUrl,
  getTeamsLinkFromUrl,
  isLandscape,
  isOnIphoneAndNotSafari,
  navigateToHomePage,
  WEB_APP_TITLE
} from './utils/AppUtils';

import { useIsMobile } from './utils/useIsMobile';
import { useSecondaryInstanceCheck } from './utils/useSecondaryInstanceCheck';
import { CallError } from './views/CallError';
import { CallScreen } from './views/CallScreen';
import { EndCall } from './views/EndCall';
import { HomeScreen } from './views/HomeScreen';
import { PageOpenInAnotherTab } from './views/PageOpenInAnotherTab';
import { UnsupportedBrowserPage } from './views/UnsupportedBrowserPage';
import { inTeams } from './utils/inTeams';
import { app, FrameContexts } from '@microsoft/teams-js';
import { useTeamsContext } from './utils/useTeamsContext';
import TabConfig from './views/TabConfig';
import { SidePanel } from './views/SidePanel';
import { NoteContainer } from './components/NoteContainer';

setLogLevel('warning');

initializeIcons();

type AppPages = 'home' | 'call' | 'endCall' | 'teamsConfig' | 'teamsSidePanel' | 'teamsMeetingStage';
942794;

const App = (): JSX.Element => {
  const [page, setPage] = useState<AppPages>('home');

  // User credentials to join a call with - these are retrieved from the server
  const [token, setToken] = useState<string>();
  const [userId, setUserId] = useState<CommunicationUserIdentifier>();
  const [userCredentialFetchError, setUserCredentialFetchError] = useState<boolean>(false);

  // Call details to join a call - these are collected from the user on the home screen
  const [callLocator, setCallLocator] = useState<CallAdapterLocator | null>(null);
  const [displayName, setDisplayName] = useState<string>('ACS Teams Guest User');
  const [initialized, setInitialized] = useState(false);
  const teamsContext = useTeamsContext(initialized);

  useEffect(() => {
    const initialize = async () => {
      try {
        console.log('App.js: initializing client SDK initialized');
        await app.initialize();
        app.notifyAppLoaded();
        app.notifySuccess();
        setInitialized(true);
      } catch (error) {
        console.error(error);
      }
    };

    if (inTeams() && !initialized) {
      console.log('App.js: initializing client SDK');
      initialize();
    }
  }, [initialized]);

  // Get Azure Communications Service token from the server
  useEffect(() => {
    (async () => {
      try {
        const { token, user } = await fetchTokenResponse();
        setToken(token);
        setUserId(user);
      } catch (e) {
        console.error(e);
        setUserCredentialFetchError(true);
      }
    })();
  }, []);

  const isMobileSession = useIsMobile();
  const isLandscapeSession = isLandscape();
  const isAppAlreadyRunningInAnotherTab = useSecondaryInstanceCheck();

  useEffect(() => {
    if (isMobileSession && isLandscapeSession) {
      console.log('ACS Calling sample: Mobile landscape view is experimental behavior');
    }
  }, [isMobileSession, isLandscapeSession]);

  if (isMobileSession && isAppAlreadyRunningInAnotherTab) {
    return <PageOpenInAnotherTab />;
  }

  const teamsAppReady = inTeams() && initialized;
  useEffect(() => {
    if (teamsAppReady && teamsContext) {
      console.log('App.js: teams app ready');
      // find context
      const frameContext = teamsContext.page.frameContext;
      console.log('App.js: frameContext', frameContext);
      if (frameContext === FrameContexts.settings) {
        setPage('teamsConfig');
      } else if (frameContext === FrameContexts.sidePanel) {
        setPage('teamsSidePanel');
      } else if (frameContext === FrameContexts.meetingStage) {
        setPage('teamsMeetingStage');
      }
    }
  }, [teamsAppReady, teamsContext]);

  const supportedBrowser = !isOnIphoneAndNotSafari();
  if (!supportedBrowser) {
    return <UnsupportedBrowserPage />;
  }

  switch (page) {
    case 'teamsConfig': {
      return <TabConfig />;
    }
    case 'teamsSidePanel': {
      return <SidePanel />;
    }
    case 'teamsMeetingStage': {
      return <NoteContainer acsLiveShareHostOptions={undefined} />;
    }
    case 'home': {
      if (inTeams() && !initialized && !teamsContext) {
        return <Spinner label={'Initializing Teams...'} ariaLive="assertive" labelPosition="top" />;
      }

      document.title = `home - ${WEB_APP_TITLE}`;
      // Show a simplified join home screen if joining an existing call
      const joiningExistingCall: boolean = !!getGroupIdFromUrl() || !!getTeamsLinkFromUrl();

      return (
        <HomeScreen
          joiningExistingCall={joiningExistingCall}
          startCallHandler={async (callDetails) => {
            setDisplayName(callDetails.displayName);
            let callLocator: CallAdapterLocator | undefined =
              callDetails.callLocator || getTeamsLinkFromUrl() || getGroupIdFromUrl();

            callLocator = callLocator || createGroupId();

            setCallLocator(callLocator);
            setPage('call');
          }}
        />
      );
    }
    case 'endCall': {
      document.title = `end call - ${WEB_APP_TITLE}`;
      return <EndCall rejoinHandler={() => setPage('call')} homeHandler={navigateToHomePage} />;
    }
    case 'call': {
      if (userCredentialFetchError) {
        document.title = `error - ${WEB_APP_TITLE}`;
        return (
          <CallError
            title="Error getting user credentials from server"
            reason="Ensure the sample server is running."
            rejoinHandler={() => setPage('call')}
            homeHandler={navigateToHomePage}
          />
        );
      }

      if (!token || !userId || !displayName || !callLocator) {
        document.title = `credentials - ${WEB_APP_TITLE}`;
        return <Spinner label={'Getting user credentials from server'} ariaLive="assertive" labelPosition="top" />;
      }
      return (
        <CallScreen
          token={token}
          userId={userId}
          displayName={displayName}
          callLocator={callLocator}
          onCallEnded={() => setPage('endCall')}
        />
      );
    }
    default:
      document.title = `error - ${WEB_APP_TITLE}`;
      return <>Invalid page</>;
  }
};

const getJoinParams = (locator: CallAdapterLocator): string => {
  if ('meetingLink' in locator) {
    return '?teamsLink=' + encodeURIComponent(locator.meetingLink);
  }

  return '?groupId=' + encodeURIComponent(locator.groupId);
};

export default App;
