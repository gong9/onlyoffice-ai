import { Injectable, Logger } from '@nestjs/common';
import * as WebSocket from 'ws';
import { Server } from 'http';
import { WSClient, WSMessage, WSMessageType } from './websocket.interface';

@Injectable()
export class WebSocketService {
  private readonly logger = new Logger(WebSocketService.name);
  private wss: WebSocket.Server;
  private clients: Map<string, WSClient> = new Map();
  private connectGroups: Map<string, string[]> = new Map(); // connectId -> [clientId1, clientId2]
  private pingInterval: NodeJS.Timeout;

  constructor() {
    // 启动心跳检测
    this.startHeartbeat();
  }

  /**
   * 初始化 WebSocket 服务器
   * @param server
   * @param path
   */
  initialize(server: Server, path = '/ws'): void {
    this.wss = new WebSocket.Server({
      server,
      path,
      perMessageDeflate: false,
    });

    this.wss.on('connection', (ws: WebSocket, request) => {
      console.log('WebSocket 连接', this.clients.size);
      this.handleConnection(ws, request);
    });

    this.logger.log(`WebSocket 服务器已启动，路径: ${path}`);
  }

  /**
   * 处理新的连接
   */
  private handleConnection(ws: WebSocket, request: any): void {
    const clientId = this.generateClientId();
    const client: WSClient = {
      id: clientId,
      ws,
      isAlive: true,
      lastPing: Date.now(),
    };

    this.clients.set(clientId, client);
    this.logger.log(
      `客户端连接: ${clientId}, 当前连接数: ${this.clients.size}`,
    );

    // 发送连接成功消息
    this.sendToClient(clientId, {
      type: WSMessageType.CONNECT,
      data: { clientId, message: '连接成功' },
      timestamp: Date.now(),
    });

    // 处理消息
    ws.on('message', (data: WebSocket.Data) => {
      this.handleMessage(clientId, data);
    });

    // 处理 pong 响应
    ws.on('pong', () => {
      if (this.clients.has(clientId)) {
        this.clients.get(clientId)!.isAlive = true;
        this.clients.get(clientId)!.lastPing = Date.now();
      }
    });

    // 处理连接关闭
    ws.on('close', (code: number, reason: Buffer) => {
      this.handleDisconnection(clientId, code, reason);
    });

    // 处理错误
    ws.on('error', (error: Error) => {
      this.logger.error(`客户端 ${clientId} 连接错误:`, error);
      this.handleDisconnection(clientId);
    });
  }

  /**
   * 处理收到的消息
   */
  private handleMessage(clientId: string, data: WebSocket.Data): void {
    try {
      const message: WSMessage = JSON.parse(data.toString());
      this.logger.debug(`收到来自 ${clientId} 的消息:`, message);

      // 更新客户端活跃状态
      if (this.clients.has(clientId)) {
        this.clients.get(clientId)!.isAlive = true;
        this.clients.get(clientId)!.lastPing = Date.now();
      }

      // 处理不同类型的消息
      switch (message.type) {
        case WSMessageType.PING:
          this.sendToClient(clientId, {
            type: WSMessageType.PONG,
            timestamp: Date.now(),
          });
          break;

        case WSMessageType.JOIN_ROOM:
          this.handleConnectId(clientId, message);
          break;

        case WSMessageType.PRIVATE_MESSAGE:
          this.handlePrivateMessage(clientId, message);
          break;

        case WSMessageType.MESSAGE:
          this.logger.log(`处理来自 ${clientId} 的业务消息:`, message.data);
          break;

        default:
          this.sendToClient(clientId, {
            type: WSMessageType.ERROR,
            data: { message: '未知的消息类型' },
            timestamp: Date.now(),
          });
      }
    } catch (error) {
      this.logger.error(`解析消息失败 from ${clientId}:`, error);
      this.sendToClient(clientId, {
        type: WSMessageType.ERROR,
        data: { message: '消息格式错误' },
        timestamp: Date.now(),
      });
    }
  }

  /**
   * 处理连接ID配对
   */
  private handleConnectId(clientId: string, message: WSMessage): void {
    const connectId = message.data?.connectId;

    if (!connectId) {
      this.sendToClient(clientId, {
        type: WSMessageType.ERROR,
        data: { message: '缺少连接ID' },
        timestamp: Date.now(),
      });
      return;
    }

    const client = this.clients.get(clientId);
    if (!client) {
      return;
    }

    // 设置客户端的连接ID
    client.metadata = { connectId };

    // 获取当前连接组
    if (!this.connectGroups.has(connectId)) {
      this.connectGroups.set(connectId, []);
    }

    const group = this.connectGroups.get(connectId)!;

    // 如果客户端已经在组里，不重复添加
    if (!group.includes(clientId)) {
      group.push(clientId);
    }

    this.logger.log(
      `客户端 ${clientId} 使用连接ID ${connectId} 配对, 当前组内人数: ${group.length}`,
    );

    // 如果正好有2个人，进行配对
    if (group.length === 2) {
      const [client1Id, client2Id] = group;

      // 通知第一个客户端
      this.sendToClient(client1Id, {
        type: WSMessageType.MESSAGE,
        data: {
          message: '配对成功！可以开始聊天了',
          partnerId: client2Id,
          status: 'paired',
          connectId: connectId,
        },
        timestamp: Date.now(),
      });

      // 通知第二个客户端
      this.sendToClient(client2Id, {
        type: WSMessageType.MESSAGE,
        data: {
          message: '配对成功！可以开始聊天了',
          partnerId: client1Id,
          status: 'paired',
          connectId: connectId,
        },
        timestamp: Date.now(),
      });
    } else if (group.length === 1) {
      // 等待配对
      this.sendToClient(clientId, {
        type: WSMessageType.MESSAGE,
        data: {
          message: '等待其他人使用相同连接ID加入...',
          status: 'waiting',
          connectId: connectId,
        },
        timestamp: Date.now(),
      });
    } else if (group.length > 2) {
      this.sendToClient(clientId, {
        type: WSMessageType.MESSAGE,
        data: {
          message: `连接ID ${connectId} 已有 ${group.length} 人，建议使用新的连接ID`,
          status: 'warning',
          connectId: connectId,
        },
        timestamp: Date.now(),
      });
    }
  }

  /**
   * 处理一对一私信
   */
  private handlePrivateMessage(fromClientId: string, message: WSMessage): void {
    const fromClient = this.clients.get(fromClientId);
    if (!fromClient || !fromClient.metadata?.connectId) {
      this.sendToClient(fromClientId, {
        type: WSMessageType.ERROR,
        data: { message: '您还没有配对，请先使用连接ID进行配对' },
        timestamp: Date.now(),
      });
      return;
    }

    const connectId = fromClient.metadata.connectId;
    const group = this.connectGroups.get(connectId);

    if (!group || group.length !== 2) {
      this.sendToClient(fromClientId, {
        type: WSMessageType.ERROR,
        data: { message: '暂无配对伙伴，无法发送消息' },
        timestamp: Date.now(),
      });
      return;
    }

    // 找到配对的伙伴
    const partnerId = group.find((id) => id !== fromClientId);

    if (!partnerId || !this.clients.has(partnerId)) {
      this.sendToClient(fromClientId, {
        type: WSMessageType.ERROR,
        data: { message: '配对伙伴已断开连接' },
        timestamp: Date.now(),
      });
      return;
    }

    // 转发消息给配对伙伴
    const forwardMessage: WSMessage = {
      type: WSMessageType.PRIVATE_MESSAGE,
      data: message.data,
      fromClientId: fromClientId,
      timestamp: Date.now(),
    };

    const success = this.sendToClient(partnerId, forwardMessage);

    // 给发送方回复发送状态
    this.sendToClient(fromClientId, {
      type: WSMessageType.MESSAGE,
      data: {
        status: success ? 'delivered' : 'failed',
        message: success ? '消息已发送' : '消息发送失败',
        toClientId: partnerId,
      },
      timestamp: Date.now(),
    });

    this.logger.log(
      `配对消息: ${fromClientId} -> ${partnerId} (连接ID: ${connectId}), 状态: ${
        success ? '成功' : '失败'
      }`,
    );
  }

  /**
   * 处理连接断开
   */
  private handleDisconnection(
    clientId: string,
    code?: number,
    reason?: Buffer,
  ): void {
    if (this.clients.has(clientId)) {
      const client = this.clients.get(clientId);

      // 从连接组中移除客户端
      if (client?.metadata?.connectId) {
        const group = this.connectGroups.get(client.metadata.connectId);
        if (group) {
          const index = group.indexOf(clientId);
          if (index > -1) {
            group.splice(index, 1);
            this.logger.log(
              `从连接组 ${client.metadata.connectId} 中移除客户端 ${clientId}, 剩余组内人数: ${group.length}`,
            );

            // 如果组内没有人了，删除整个组
            if (group.length === 0) {
              this.connectGroups.delete(client.metadata.connectId);
              this.logger.log(`连接组 ${client.metadata.connectId} 已删除`);
            } else {
              // 如果组内还有其他人，通知他们伙伴已断开
              for (const remainingClientId of group) {
                this.sendToClient(remainingClientId, {
                  type: WSMessageType.MESSAGE,
                  data: {
                    message: '你的配对伙伴已断开连接',
                    status: 'partner_disconnected',
                    connectId: client.metadata.connectId,
                  },
                  timestamp: Date.now(),
                });
              }
            }
          }
        }
      }

      this.clients.delete(clientId);
      this.logger.log(
        `客户端断开连接: ${clientId}, 断开码: ${code}, 剩余连接数: ${this.clients.size}`,
      );
    }
  }

  /**
   * 发送消息给指定客户端
   */
  public sendToClient(clientId: string, message: WSMessage): boolean {
    const client = this.clients.get(clientId);
    if (!client || client.ws.readyState !== WebSocket.OPEN) {
      this.logger.warn(`客户端 ${clientId} 不存在或连接已关闭`);
      return false;
    }

    try {
      client.ws.send(JSON.stringify(message));
      return true;
    } catch (error) {
      this.logger.error(`发送消息给客户端 ${clientId} 失败:`, error);
      this.handleDisconnection(clientId);
      return false;
    }
  }

  /**
   * 生成客户端ID
   */
  private generateClientId(): string {
    return `client_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * 启动心跳检测
   */
  private startHeartbeat(): void {
    if (this.pingInterval) {
      clearInterval(this.pingInterval);
    }

    this.pingInterval = setInterval(() => {
      const now = Date.now();
      const timeout = 30000;

      for (const [clientId, client] of this.clients) {
        const lastPingTime = client.lastPing || 0;

        if (!client.isAlive || now - lastPingTime > timeout) {
          this.logger.warn(`客户端 ${clientId} 心跳超时，断开连接`);

          if (
            client.ws.readyState === WebSocket.OPEN ||
            client.ws.readyState === WebSocket.CONNECTING
          ) {
            client.ws.terminate();
          }

          this.handleDisconnection(clientId);
        } else {
          // 只有当连接正常时才发送 ping 和设置 isAlive
          if (client.ws.readyState === WebSocket.OPEN) {
            client.isAlive = false;
            client.ws.ping();
          }
        }
      }
    }, 10000); // 每10秒检查一次
  }

  /**
   * 关闭 WebSocket 服务
   */
  public close(): void {
    if (this.pingInterval) {
      clearInterval(this.pingInterval);
    }

    if (this.wss) {
      this.wss.close();
      this.logger.log('WebSocket 服务器已关闭');
    }

    // 关闭所有客户端连接
    for (const [clientId, client] of this.clients) {
      client.ws.close();
    }
    this.clients.clear();
  }
}
