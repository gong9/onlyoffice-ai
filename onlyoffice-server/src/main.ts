import { NestFactory } from '@nestjs/core';
import { ValidationPipe, VersioningType } from '@nestjs/common';
import { SwaggerModule, DocumentBuilder } from '@nestjs/swagger';
import { AppModule } from './app.module';
import { ResponseInterceptor } from './shared/interceptors/response.interceptor';
import { LoggingInterceptor } from './shared/interceptors/logger.interceptor';
import { WebSocketService } from './websocket/websocket.service';

async function bootstrap() {
  const app = await NestFactory.create(AppModule, {
    cors: {
      origin: true,
      credentials: true,
      methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
      allowedHeaders: ['Content-Type', 'Authorization'],
    },
  });

  app.setGlobalPrefix(process.env.API_PREFIX);

  app.useGlobalPipes(
    new ValidationPipe({
      // 跳过验证对象中值为 null 或 undefined 的属性的验证。完整配置文档参见：https://docs.nestjs.cn/9/techniques?id=%e9%aa%8c%e8%af%81
      skipNullProperties: true,
      stopAtFirstError: true,
      transform: true,
    }),
  );

  app.useGlobalInterceptors(
    new ResponseInterceptor(),
    new LoggingInterceptor(),
  );

  app.enableVersioning({
    type: VersioningType.URI,
    defaultVersion: '1',
  });

  const isTest = process.env.NODE_ENV === 'test';
  if (isTest) {
    const options = new DocumentBuilder()
      .setTitle('Demo Onlyoffice 接口文档')
      .setVersion('1.0.0')
      .addTag('Onlyoffice')
      .addTag('Document')
      .addTag('WebSocket')
      .build();
    const document = SwaggerModule.createDocument(app, options);
    SwaggerModule.setup('docs', app, document);
  }

  // 启动 HTTP 服务器
  await app.listen(process.env.PORT);

  // 获取底层 HTTP 服务器实例
  const httpServer = app.getHttpServer();

  // 获取 WebSocket 服务并初始化
  const websocketService = app.get(WebSocketService);
  websocketService.initialize(httpServer, `${process.env.API_PREFIX}/v1/ws`);

  console.log(`应用已启动在端口 ${process.env.PORT}`);
  console.log(
    `WebSocket 服务已启动，连接地址: ws://localhost:${process.env.PORT}${process.env.API_PREFIX}/v1/ws`,
  );
}
bootstrap();
