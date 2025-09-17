import { Controller } from '@nestjs/common';
import { ApiTags } from '@nestjs/swagger';
import { WebSocketService } from './websocket.service';

@ApiTags('WebSocket')
@Controller('websocket')
export class WebSocketController {
  constructor(private readonly websocketService: WebSocketService) {}
}
