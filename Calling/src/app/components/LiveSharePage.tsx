// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Stack } from '@fluentui/react';
import { Spinner } from '@fluentui/react';
import { IFluidContainer } from 'fluid-framework';
import React from 'react';
import { FC, ReactNode, useMemo } from 'react';

export const LiveSharePage: FC<{
  children: ReactNode;
  started: boolean;
  container?: IFluidContainer;
}> = ({ children, started, container }) => {
  const loadText = useMemo(() => {
    if (!container) {
      return 'Joining Live Share session...';
    }
    if (!started) {
      return 'Starting sync...';
    }
    return undefined;
  }, [container, started]);

  return (
    <>
      {loadText && (
        <Stack>
          <Spinner label={loadText} ariaLive="assertive" labelPosition="top" />
        </Stack>
      )}
      <div style={{ visibility: loadText ? 'hidden' : undefined }}>{children}</div>
    </>
  );
};
