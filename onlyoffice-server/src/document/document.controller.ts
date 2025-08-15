import {
  Controller,
  Get,
  Post,
  Body,
  Query,
  UseInterceptors,
  UploadedFile,
  BadRequestException,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { ApiTags, ApiOperation, ApiConsumes, ApiBody } from '@nestjs/swagger';
import { diskStorage } from 'multer';
import { extname, join } from 'path';
import { DocumentService } from './document.service';
import {
  DocumentInfoDto,
  DocumentForceSaveDto,
  UploadFileDto,
  FileUploadResponseDto,
} from './document.dto';

@ApiTags('文档管理')
@Controller({
  path: 'document',
  version: '1',
})
export class DocumentController {
  constructor(private readonly documentService: DocumentService) {}

  @Get('/info')
  @ApiOperation({ summary: '获取文档信息' })
  async documentInfo(@Query() query: DocumentInfoDto) {
    return await this.documentService.documentInfo(query);
  }

  @Post('/forceSave')
  @ApiOperation({ summary: '强制保存文档' })
  async forceSave(@Body() body: DocumentForceSaveDto) {
    return await this.documentService.forceSave(body);
  }

  @Post('/upload')
  @ApiOperation({ summary: '上传文档文件' })
  @ApiConsumes('multipart/form-data')
  @ApiBody({
    description: '文件上传',
    type: 'multipart/form-data',
    schema: {
      type: 'object',
      properties: {
        file: {
          type: 'string',
          format: 'binary',
          description: '要上传的文档文件',
        },
        fileName: {
          type: 'string',
          description: '自定义文件名（可选）',
        },
        description: {
          type: 'string',
          description: '文件描述（可选）',
        },
      },
      required: ['file'],
    },
  })
  @UseInterceptors(
    FileInterceptor('file', {
      storage: diskStorage({
        destination: './static',
        filename: (req, file, callback) => {
          // 生成唯一文件名
          // eslint-disable-next-line prettier/prettier
          const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
          const ext = extname(file.originalname);
          callback(null, `upload-${uniqueSuffix}${ext}`);
        },
      }),
      fileFilter: (req, file, callback) => {
        // 只允许 Office 文档格式
        const allowedMimes = [
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
          'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
          'application/msword', // .doc
          'application/vnd.ms-excel', // .xls
          'application/vnd.ms-powerpoint', // .ppt
        ];

        if (allowedMimes.includes(file.mimetype)) {
          callback(null, true);
        } else {
          // eslint-disable-next-line prettier/prettier
          callback(new BadRequestException('只支持 Office 文档格式 (.docx, .xlsx, .pptx, .doc, .xls, .ppt)'), false);
        }
      },
      limits: {
        fileSize: 50 * 1024 * 1024, // 50MB 限制
      },
    }),
  )
  async uploadFile(
    @UploadedFile() file: Express.Multer.File,
    @Body() uploadDto: UploadFileDto,
  ): Promise<FileUploadResponseDto> {
    return await this.documentService.uploadFile(file, uploadDto);
  }

  @Get('/list')
  @ApiOperation({ summary: '获取文档列表' })
  async getDocumentList() {
    return await this.documentService.getDocumentList();
  }
}
