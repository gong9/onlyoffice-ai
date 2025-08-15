import { ConfigService } from '@nestjs/config';
import { Injectable, HttpStatus, HttpException } from '@nestjs/common';
import { OnlyofficeService } from '../onlyoffice/onlyoffice.service';
import {
  DocumentForceSaveDto,
  DocumentInfoDto,
  UploadFileDto,
  FileUploadResponseDto,
} from './document.dto';
import { DocumentForceSave, DocumentInfo } from './document.entity';
import { readdirSync, statSync } from 'fs';
import { join, extname } from 'path';

@Injectable()
export class DocumentService {
  constructor(
    private readonly config: ConfigService,
    private readonly onlyofficeService: OnlyofficeService,
  ) {}

  async forceSave(body: DocumentForceSaveDto): Promise<DocumentForceSave> {
    // 1、保存业务数据
    // 2、调用 Onlyoffice 的强制保存，实际业务中可能还有更多的业务操作，可根据实际情况删改
    const { id: userdata, key, useJwtEncrypt } = body;
    const data = await this.onlyofficeService.forceSave({
      key,
      // 将业务参数传给 Onlyoffice 服务，当回调里面存在多个请求时，标识符将有助于区分特定请求
      userdata,
      useJwtEncrypt,
    });
    // 保存成功
    if (data.error === 0) {
      return null;
    }
    throw new HttpException(data, HttpStatus.OK);
  }

  async documentInfo(query: DocumentInfoDto): Promise<DocumentInfo> {
    const editorConfig = this.onlyofficeService.editorDefaultConfig();
    // 根据文件扩展名确定文件类型
    const fileExt = extname(query.key).toLowerCase().substring(1);
    const fileType = fileExt || 'docx';

    // 根据扩展名确定文档类型
    let documentType = 'word';
    if (['xlsx', 'xls'].includes(fileExt)) {
      documentType = 'cell';
    } else if (['pptx', 'ppt'].includes(fileExt)) {
      documentType = 'slide';
    }

    // 设置文档类型
    editorConfig.documentType = documentType;

    // 添加文档
    editorConfig.document = {
      ...editorConfig.document,
      fileType: fileType,
      key: query.key,
      url: `${this.config.get('domain')}/static/${query.key}`,
      title: query.key,
    };
    // 添加用户信息
    editorConfig.editorConfig.user = {
      group: '技术部',
      id: 'wytxer',
      name: 'gz',
    };
    // 添加插件配置

    editorConfig.editorConfig.plugins = {
      autostart: ['asc.{11700c35-1fdb-4e37-9edb-b31637139601}'],
      pluginsData: [
        `http://localhost:3000/static/plugins/plugin-hello/config.json`,
      ],
    };

    // 加密编辑器参数
    if (query.useJwtEncrypt === 'y') {
      this.onlyofficeService.signJwt(editorConfig);
    }
    return {
      id: 1,
      remarks: '业务字段',
      editorConfig,
    };
  }

  async excelInfo(query: DocumentInfoDto): Promise<DocumentInfo> {
    const editorConfig = this.onlyofficeService.editorDefaultConfig();
    // 添加文档
    editorConfig.document = {
      ...editorConfig.document,
      fileType: 'xlsx',
      key: query.key,
      url: `${this.config.get('domain')}/static/${query.key}`,
      title: '测试表格.xlsx',
    };
    // 修改文档宽度
    editorConfig.width = '100%';
    // 修改编辑器类型
    editorConfig.documentType = 'cell';
    // 添加用户信息
    editorConfig.editorConfig.user = {
      group: '技术部',
      id: 'wytxer',
      name: 'gz',
    };
    // 加密编辑器参数
    if (query.useJwtEncrypt === 'y') {
      this.onlyofficeService.signJwt(editorConfig);
    }
    return {
      id: 1,
      remarks: '业务字段',
      editorConfig,
    };
  }

  /**
   * 上传文件
   */
  async uploadFile(
    file: Express.Multer.File,
    uploadDto: UploadFileDto,
  ): Promise<FileUploadResponseDto> {
    try {
      if (!file) {
        return {
          code: 400,
          message: '请选择要上传的文件',
        };
      }

      // 确定文件类型
      const fileExtension = extname(file.originalname).toLowerCase();

      // eslint-disable-next-line prettier/prettier
      const supportedExtensions = [
        '.docx',
        '.xlsx',
        '.pptx',
        '.doc',
        '.xls',
        '.ppt',
      ];
      if (!supportedExtensions.includes(fileExtension)) {
        return {
          code: 400,
          message: '不支持的文件格式，请上传 Office 文档文件',
        };
      }

      // 生成文件唯一标识（使用文件名作为 key）
      const fileKey = file.filename;
      return {
        code: 200,
        message: '文件上传成功',
        data: {
          key: fileKey,
          fileName: uploadDto.fileName || file.originalname,
          fileSize: file.size,
          filePath: file.path,
          uploadTime: new Date().toISOString(),
        },
      };
    } catch (error) {
      return {
        code: 500,
        message: '文件上传失败：' + error.message,
      };
    }
  }

  /**
   * 获取文档列表
   */
  async getDocumentList() {
    try {
      const staticPath = this.config.get('staticPath');
      const files = readdirSync(staticPath);

      const documentFiles = files
        .filter((file) => {
          const ext = extname(file).toLowerCase();
          return ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt'].includes(
            ext,
          );
        })
        .map((file) => {
          const filePath = join(staticPath, file);
          const stats = statSync(filePath);
          const ext = extname(file).toLowerCase();

          // 根据扩展名确定文档类型
          let documentType = 'word';
          if (['.xlsx', '.xls'].includes(ext)) {
            documentType = 'cell';
          } else if (['.pptx', '.ppt'].includes(ext)) {
            documentType = 'slide';
          }

          return {
            key: file,
            fileName: file,
            fileSize: stats.size,
            documentType,
            uploadTime: stats.mtime.toISOString(),
            url: `${this.config.get('domain')}/static/${file}`,
          };
        })
        .sort(
          (a, b) =>
            new Date(b.uploadTime).getTime() - new Date(a.uploadTime).getTime(),
        );

      return {
        code: 200,
        message: '获取文档列表成功',
        data: {
          total: documentFiles.length,
          files: documentFiles,
        },
      };
    } catch (error) {
      return {
        code: 500,
        message: '获取文档列表失败：' + error.message,
      };
    }
  }
}
