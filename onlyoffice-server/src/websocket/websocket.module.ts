import { Module } from '@nestjs/common';
import { WebSocketService } from './websocket.service';
import { WebSocketController } from './websocket.controller';

@Module({
  controllers: [WebSocketController],
  providers: [WebSocketService],
  exports: [WebSocketService],
})
export class WebSocketModule {}
