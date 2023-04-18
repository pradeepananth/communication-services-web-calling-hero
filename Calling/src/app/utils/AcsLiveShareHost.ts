import {
  ContainerState,
  IFluidContainerInfo,
  IFluidTenantInfo,
  ILiveShareHost,
  INtpTimeInfo,
  UserMeetingRole
} from '@microsoft/live-share';
import { CallAdapter } from '@azure/communication-react';
const LiveShareRoutePrefix = '/livesync/v1/acs';
const LiveShareBaseUrl = 'https://teams.microsoft.com/api/platform';
const GetNtpTimeRoute = 'getNTPTime';
const GetFluidTenantInfoRoute = 'fluid/tenantInfo/get';
const RegisterClientRolesRoute = 'clientRoles/register';
const ClientRolesGetRoute = 'clientRoles/get';
const FluidTokenGetRoute = 'fluid/token/get';
const FluidContainerGetRoute = 'fluid/container/get';
const FluidContainerSetRoute = 'fluid/container/set';

export class AcsLiveShareHost implements ILiveShareHost {
  private constructor(
    private readonly acsToken: string,
    private readonly callAdapter: CallAdapter,
    private readonly meetingJoinUrl: string
  ) {
    if (!callAdapter.getState().isTeamsCall) {
      throw new Error('only teams calls are supported');
    }
  }

  public static create(options: AcsLiveShareHostOptions): ILiveShareHost {
    return new AcsLiveShareHost(options.acsTokenProvider(), options.callAdapter, options.teamsMeetingJoinUrl);
  }

  async getClientRoles(clientId: string): Promise<UserMeetingRole[] | undefined> {
    const request = this.constructBaseRequest() as FluidClientRolesInput;
    request.clientId = clientId;
    const response = await fetch(`${LiveShareBaseUrl}/${LiveShareRoutePrefix}/${ClientRolesGetRoute}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `SkypeToken ${this.acsToken}`
      },
      body: JSON.stringify(request)
    });
    const data = await response.json();
    return data.roles;
  }

  async getFluidContainerId(): Promise<IFluidContainerInfo> {
    const request = this.constructBaseRequest() as FluidGetContainerIdInput;
    const response = await fetch(`${LiveShareBaseUrl}/${LiveShareRoutePrefix}/${FluidContainerGetRoute}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `SkypeToken ${this.acsToken}`
      },
      body: JSON.stringify(request)
    });
    const data = await response.json();
    return data;
  }

  async getFluidTenantInfo(): Promise<IFluidTenantInfo> {
    const request = this.constructBaseRequest() as FluidTenantInfoInput;
    request.expiresAt = 0;
    const response = await fetch(`${LiveShareBaseUrl}/${LiveShareRoutePrefix}/${GetFluidTenantInfoRoute}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `SkypeToken ${this.acsToken}`
      },
      body: JSON.stringify(request)
    });

    const data = await response.json();
    return data.broadcaster.frsTenantInfo;
  }

  async getFluidToken(containerId?: string): Promise<string> {
    const request = this.constructBaseRequest() as FluidGetTokenInput;
    request.containerId = containerId;
    const response = await fetch(`${LiveShareBaseUrl}/${LiveShareRoutePrefix}/${FluidTokenGetRoute}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `SkypeToken ${this.acsToken}`
      },
      body: JSON.stringify(request)
    });
    const data = await response.json();
    return data.token;
  }

  async getNtpTime(): Promise<INtpTimeInfo> {
    const response = await fetch(`${LiveShareBaseUrl}/${LiveShareRoutePrefix}/${GetNtpTimeRoute}`, {
      method: 'GET'
    });
    const data = await response.json();
    return data;
  }

  async registerClientId(clientId: string): Promise<UserMeetingRole[]> {
    const request = this.constructBaseRequest() as FluidClientRolesInput;
    request.clientId = clientId;
    const response = await fetch(`${LiveShareBaseUrl}/${LiveShareRoutePrefix}/${RegisterClientRolesRoute}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `SkypeToken ${this.acsToken}`
      },
      body: JSON.stringify(request)
    });
    const data = await response.json();
    return data;
  }

  async setFluidContainerId(containerId: string): Promise<IFluidContainerInfo> {
    const request = this.constructBaseRequest() as FluidSetContainerIdInput;
    request.containerId = containerId;
    const response = await fetch(`${LiveShareBaseUrl}/${LiveShareRoutePrefix}/${FluidContainerSetRoute}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `SkypeToken ${this.acsToken}`
      },
      body: JSON.stringify(request)
    });
    const data = await response.json();
    return data;
  }

  private constructBaseRequest(): LiveShareRequestBase {
    const userId = this.callAdapter.getState().userId;
    if (userId.kind !== 'communicationUser') {
      throw new Error(`unsupported user id ${userId.kind}`);
    }
    const originUri = window.location.href;
    return {
      originUri,
      teamsContextType: TeamsCollabContextType.MeetingJoinUrl,
      teamsContext: {
        meetingJoinUrl: this.meetingJoinUrl,
        skypeMri: userId.communicationUserId
      }
    };
  }
}

export interface AcsLiveShareHostOptions {
  callAdapter: CallAdapter;
  teamsMeetingJoinUrl: string;
  acsTokenProvider: () => string;
}

interface FluidCollabTenantInfo {
  broadcaster: BroadcasterInfo;
}

interface BroadcasterInfo {
  type: string;
  frsTenantInfo: FluidTenantInfo;
}

interface FluidTenantInfo {
  tenantId: string;
  ordererEndpoint: string;
  storageEndpoint: string;
  serviceEndpoint: string;
}

interface FluidTenantInfoInput {
  appId?: string;
  originUri: string;
  teamsContextType: TeamsCollabContextType;
  teamsContext: TeamsContext;
  expiresAt: number;
}

interface FluidGetContainerIdInput extends LiveShareRequestBase {}

interface TeamsContext {
  meetingJoinUrl?: string;
  skypeMri?: string;
}

interface FluidSetContainerIdInput extends LiveShareRequestBase {
  containerId: string;
}

interface FluidContainerInfo {
  containerState: ContainerState;
  shouldCreate: boolean;
  containerId: string;
  retryAfter: number;
}

interface FluidClientRolesInput extends LiveShareRequestBase {
  clientId: string;
}

interface FluidGetTokenInput {
  appId?: string;
  originUri: string;
  teamsContextType: TeamsCollabContextType;
  teamsContext: TeamsContext;
  containerId?: string;
  // TODO: these are not used on server side    // userId?: string;    // userName?: string;
}

interface User {
  mri: string;
}

enum TeamsCollabContextType {
  MeetingJoinUrl = 1,
  GroupChatId
}

interface LiveShareRequestBase {
  appId?: string;
  originUri: string;
  teamsContextType: TeamsCollabContextType;
  teamsContext: TeamsContext;
}
