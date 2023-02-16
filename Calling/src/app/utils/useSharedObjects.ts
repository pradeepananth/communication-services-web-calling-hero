/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { LiveShareClient, TestLiveShareHost } from '@microsoft/live-share';
import { ContainerSchema, IFluidContainer, SharedMap } from 'fluid-framework';
import { useEffect, useState } from 'react';
import { LivePresence } from '@microsoft/live-share';
import { AzureConnectionConfig, AzureContainerServices } from '@fluidframework/azure-client';
import { LiveShareHost } from '@microsoft/teams-js';
import { AcsLiveShareHost, AcsLiveShareHostOptions } from './AcsLiveShareHost';
import { inTeams } from './inTeams';

/**
 * Hook that creates/loads the apps shared objects.
 *
 * @remarks
 * This is an application specific hook that defines the fluid schema of Distributed Data Structures (DDS)
 * used by the app and passes that schema to the `LiveShareClient` to create/load your Fluid container.
 *
 * @returns Shared objects managed by the apps fluid container.
 */
export function useSharedObjects(acsLiveShareHostOptions: AcsLiveShareHostOptions | undefined) {
  const [results, setResults] = useState<{
    container: IFluidContainer;
    services: AzureContainerServices;
    created: boolean;
  }>();
  const [error, setError] = useState();

  useEffect(() => {
    console.log('useSharedObjects: starting');

    try {
      // Enable debugger
      window.localStorage.debug = 'fluid:*';
    } catch (error) {
      // Some users or anonymous modes in browsers disable local storage
      console.error(error);
    }

    // Define container schema
    const schema: ContainerSchema = {
      initialObjects: {
        presence: LivePresence,
        notesMap: SharedMap
      }
    };

    let isLocal = false;
    // Create live share host
    // Uncomment below line for local testing
    // isLocal = true;
    const host = isLocal
      ? TestLiveShareHost.create()
      : inTeams()
      ? LiveShareHost.create()
      : AcsLiveShareHost.create(acsLiveShareHostOptions!);

    // Create the client, join container, and set results
    console.log('useSharedObjects: joining container');
    const client = new LiveShareClient(host);
    client
      .joinContainer(schema)
      .then((results) => {
        console.log('useSharedObjects: joined container');
        setResults(results);
      })
      .catch((err) => setError(err));
  }, []);

  const container = results?.container;
  const initialObjects = container?.initialObjects;
  return {
    presence: initialObjects?.presence as LivePresence | undefined,
    notesMap: initialObjects?.notesMap as SharedMap | undefined,
    container,
    error,
    services: results?.services
  };
}
