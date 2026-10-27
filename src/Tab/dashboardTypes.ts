export type CurrentUser = {
  userId: number;
  email: string | null;
  displayName: string | null;
  firstName?: string | null;
  lastName?: string | null;
  hasPassword?: boolean;
  m365Upn?: string | null;
  authProvider?: string | null;
  credits: number;
  subscriptionPlan?: string | null;
  activationCode?: string | null;
  clientIsActive: boolean;
};

export type Issue = {
  field: "signal_bandwidth_status" | "signal_system_status" | "signal_network_status" | string;
  currentValue: string;
  affectedParticipant: string;
  severity: string;
  recommendation: string;
};

export type SignalScoreTrendPoint = {
  timestampUtc: string;
  averageScore: number;
};

export type SignalScoreTrendResponse = {
  userId: number;
  points: SignalScoreTrendPoint[];
};

export type TelemetryRecord = Record<string, string | number | boolean | null>;

