import { IsOptional } from 'class-validator';

export class WSMessageDto {
  @IsOptional()
  type?: string;

  @IsOptional()
  data?: any;

  @IsOptional()
  timestamp?: number;
}
