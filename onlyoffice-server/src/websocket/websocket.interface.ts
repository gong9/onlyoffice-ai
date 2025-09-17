export interface WSMessage {
  type: string;
  data?: any;
  timestamp?: number;
  fromClientId?: string;
  toClientId?: string;
  roomId?: string;
}

export interface WSClient {
  id: string;
  ws: any;
  isAlive: boolean;
  lastPing?: number;
  metadata?: any;
}

export enum WSMessageType {
  PING = 'ping',
  PONG = 'pong',
  CONNECT = 'connect',
  JOIN_ROOM = 'join_room',
  PRIVATE_MESSAGE = 'private_message',
  MESSAGE = 'message',
  ERROR = 'error',
}
