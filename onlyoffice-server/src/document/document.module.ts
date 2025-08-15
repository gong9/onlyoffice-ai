import { Module } from '@nestjs/common';
import { MulterModule } from '@nestjs/platform-express';
import { ConfigModule, ConfigService } from '@nestjs/config';

import { DocumentController } from './document.controller';
import { DocumentService } from './document.service';
import { OnlyofficeService } from '../onlyoffice/onlyoffice.service';

@Module({
  imports: [
    ConfigModule,
    MulterModule.registerAsync({
      imports: [ConfigModule],
      useFactory: async (configService: ConfigService) => ({
        dest: configService.get<string>('staticPath'),
      }),
      inject: [ConfigService],
    }),
  ],
  controllers: [DocumentController],
  providers: [DocumentService, OnlyofficeService],
})
export class DocumentModule {}
