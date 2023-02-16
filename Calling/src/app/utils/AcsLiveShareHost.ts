import {
  ContainerState,
  IFluidContainerInfo,
  IFluidTenantInfo,
  ILiveShareHost,
  INtpTimeInfo,
  UserMeetingRole
} from '@microsoft/live-share';
import { AxiosInstance, CreateAxiosDefaults, default as axios } from 'axios';
import { CallAdapter } from '@azure/communication-react';
const LiveShareRoutePrefix = '/livesync/v1/acs';
const GetNtpTimeRoute = 'getNTPTime';
const GetFluidTenantInfoRoute = 'fluid/tenantInfo/get';
const RegisterClientRolesRoute = 'clientRoles/register';
const ClientRolesGetRoute = 'clientRoles/get';
const FluidTokenGetRoute = 'fluid/token/get';
const FluidContainerGetRoute = 'fluid/container/get';
const FluidContainerSetRoute = 'fluid/container/set';
export class AcsLiveShareHost implements ILiveShareHost {
  private constructor(
    private readonly axios: AxiosInstance,
    private readonly callAdapter: CallAdapter,
    private readonly meetingJoinUrl: string
  ) {
    if (!callAdapter.getState().isTeamsCall) {
      throw new Error('only teams calls are supported');
    }
  }
  async getClientRoles(clientId: string): Promise<UserMeetingRole[] | undefined> {
    const request = this.constructBaseRequest() as FluidClientRolesInput;
    request.clientId = clientId;
    const response = await this.axios.post<UserMeetingRole[]>(
      `${LiveShareRoutePrefix}/${ClientRolesGetRoute}`,
      request
    );
    return response.data;
  }
  async getFluidContainerId(): Promise<IFluidContainerInfo> {
    const request = this.constructBaseRequest() as FluidGetContainerIdInput;
    const response = await this.axios.post<IFluidContainerInfo>(
      `${LiveShareRoutePrefix}/${FluidContainerGetRoute}`,
      request
    );
    return response.data;
  }
  async getFluidTenantInfo(): Promise<IFluidTenantInfo> {
    const request = this.constructBaseRequest() as FluidTenantInfoInput;
    request.expiresAt = 0;
    const response = await this.axios.post<FluidCollabTenantInfo>(
      `${LiveShareRoutePrefix}/${GetFluidTenantInfoRoute}`,
      request
    );
    return response.data.broadcaster.frsTenantInfo;
  }
  async getFluidToken(containerId?: string): Promise<string> {
    const request = this.constructBaseRequest() as FluidGetTokenInput;
    request.containerId = containerId;
    const response = await this.axios.post<any>(`${LiveShareRoutePrefix}/${FluidTokenGetRoute}`, request);
    return response.data.token;
  }
  async getNtpTime(): Promise<INtpTimeInfo> {
    const response = await this.axios.get<INtpTimeInfo>(`${LiveShareRoutePrefix}/${GetNtpTimeRoute}`);
    return response.data;
  }
  async registerClientId(clientId: string): Promise<UserMeetingRole[]> {
    const request = this.constructBaseRequest() as FluidClientRolesInput;
    request.clientId = clientId;
    const response = await this.axios.post<UserMeetingRole[]>(
      `${LiveShareRoutePrefix}/${RegisterClientRolesRoute}`,
      request
    );
    return response.data;
  }
  async setFluidContainerId(containerId: string): Promise<IFluidContainerInfo> {
    const request = this.constructBaseRequest() as FluidSetContainerIdInput;
    request.containerId = containerId;
    const response = await this.axios.post(`${LiveShareRoutePrefix}/${FluidContainerSetRoute}`, request);
    return response.data;
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
  public static create(options: AcsLiveShareHostOptions): ILiveShareHost {
    const axiosDefaults: CreateAxiosDefaults<any> = {
      baseURL: 'https://teams.microsoft.com/api/platform'
    };
    const axiosConfig = { ...axiosDefaults, ...options?.axiosOptions };
    const axiosInstance = axios.create(axiosConfig);
    axiosInstance.interceptors.request.use((config) => {
      const { acsTokenProvider } = options;
      const token = acsTokenProvider();
      config.headers.set('Authorization', `SkypeToken ${token}`, true);
      return config;
    });
    return new AcsLiveShareHost(axiosInstance, options.callAdapter, options.teamsMeetingJoinUrl);
  }
}
export interface AcsLiveShareHostOptions {
  callAdapter: CallAdapter;
  teamsMeetingJoinUrl: string;
  acsTokenProvider: () => string;
  axiosOptions?: CreateAxiosDefaults<any>;
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
