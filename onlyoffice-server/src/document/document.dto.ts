import {
  IsString,
  IsNumber,
  IsNotEmpty,
  IsIn,
  IsOptional,
} from 'class-validator';

/**
 * 文档强制保存请求参数
 */
export class DocumentForceSaveDto {
  /**
   * 业务 id
   */
  @IsNumber()
  id: string;

  /**
   * 文档标识符
   */
  @IsString()
  key: string;

  /**
   * 用户数据
   */
  @IsString()
  @IsOptional()
  userdata?: string;

  /**
   * 使用 JWT 加密文档参数，默认不加密，需要配合 Onlyoffice 的 secret 配置使用。
   */
  @IsString()
  @IsIn(['y', 'n'])
  @IsOptional()
  useJwtEncrypt?: string = 'n';
}

/**
 * 获取文档信息请求参数
 */
export class DocumentInfoDto {
  /**
   * 文档标识符
   */
  @IsString()
  key: string;

  /**
   * 是否使用 JWT 加密
   */
  @IsString()
  @IsOptional()
  useJwtEncrypt?: string = 'n';

  /**
   * 使用插件。默认不返回插件配置
   */
  @IsString()
  @IsIn(['y', 'n'])
  @IsOptional()
  usePlugin?: string = 'n';
}

export class UploadFileDto {
  /**
   * 文件名
   */
  @IsString()
  @IsOptional()
  fileName?: string;

  /**
   * 文件描述
   */
  @IsString()
  @IsOptional()
  description?: string;
}

export class FileUploadResponseDto {
  /**
   * 上传状态码
   */
  code: number;

  /**
   * 响应消息
   */
  message: string;

  /**
   * 文件信息
   */
  data?: {
    /**
     * 文件唯一标识
     */
    key: string;
    /**
     * 文件名
     */
    fileName: string;
    /**
     * 文件大小 (字节)
     */
    fileSize: number;
    /**
     * 文件路径
     */
    filePath: string;
    /**
     * 上传时间
     */
    uploadTime: string;
  };
}
