/**
 * Input: 日志消息与上下文
 * Output: 结构化日志输出
 * Pos: MCP 日志系统模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 日志工具 - 老王的日志系统
 * 写日志比写代码还规范，报错了至少知道是哪个SB搞出来的
 */

import winston from 'winston';
import path from 'path';

/**
 * 日志级别枚举
 */
export enum LogLevel {
  ERROR = 'error',
  WARN = 'warn',
  INFO = 'info',
  DEBUG = 'debug',
}

/**
 * 自定义日志格式 - 老王风格
 */
const wangFormat = winston.format.printf(({ level, message, timestamp, ...meta }) => {
  const metaStr = Object.keys(meta).length ? ` | ${JSON.stringify(meta)}` : '';
  return `[${timestamp}] [${level.toUpperCase()}] ${message}${metaStr}`;
});

/**
 * 解析布尔环境变量
 */
const isEnvTrue = (value?: string): boolean => {
  if (!value) return false;
  return ['1', 'true', 'yes', 'on'].includes(value.toLowerCase());
};

/**
 * 创建Logger实例
 */
const createLogger = (name: string): winston.Logger => {
  const homeDir = process.env.HOME || process.env.USERPROFILE || require('os').homedir();
  const logDir = path.join(homeDir, '.wps-office-mcp', 'logs');
  // MCP 默认走 stdio 协议，stdout 必须保持纯净；仅在显式开启时才输出 Console 日志
  const enableConsoleLog = isEnvTrue(process.env.MCP_CONSOLE_LOG);
  const transports: winston.transport[] = [
    // 文件输出 - 错误单独记录，出问题好找
    new winston.transports.File({
      filename: path.join(logDir, 'error.log'),
      level: 'error',
    }),
    // 所有日志
    new winston.transports.File({
      filename: path.join(logDir, 'combined.log'),
    }),
  ];

  if (enableConsoleLog) {
    transports.unshift(
      new winston.transports.Console({
        format: winston.format.combine(
          winston.format.colorize({ all: true }),
          winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
          wangFormat
        ),
      })
    );
  }

  return winston.createLogger({
    level: process.env.LOG_LEVEL || 'info',
    format: winston.format.combine(
      winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
      winston.format.errors({ stack: true }),
      wangFormat
    ),
    defaultMeta: { service: name },
    transports,
  });
};

/**
 * 主Logger - MCP Server专用
 */
export const logger = createLogger('wps-mcp');

/**
 * 创建子Logger - 各模块用自己的
 */
export const createChildLogger = (moduleName: string): winston.Logger => {
  return logger.child({ module: moduleName });
};

/**
 * 快捷日志函数 - 懒人专用
 */
export const log = {
  /**
   * 记录信息日志
   */
  info: (message: string, meta?: Record<string, unknown>): void => {
    logger.info(message, meta);
  },

  /**
   * 记录警告日志
   */
  warn: (message: string, meta?: Record<string, unknown>): void => {
    logger.warn(message, meta);
  },

  /**
   * 记录错误日志 - 艹，又出错了
   */
  error: (message: string, error?: Error | unknown, meta?: Record<string, unknown>): void => {
    if (error instanceof Error) {
      logger.error(message, { ...meta, error: error.message, stack: error.stack });
    } else {
      logger.error(message, { ...meta, error });
    }
  },

  /**
   * 记录调试日志 - 排查问题用的
   */
  debug: (message: string, meta?: Record<string, unknown>): void => {
    logger.debug(message, meta);
  },
};

/**
 * 请求日志中间件风格的函数
 */
export const logRequest = (
  method: string,
  params?: Record<string, unknown>
): void => {
  logger.info(`[REQUEST] ${method}`, { params });
};

/**
 * 响应日志
 */
export const logResponse = (
  method: string,
  success: boolean,
  duration: number,
  data?: unknown
): void => {
  const level = success ? 'info' : 'error';
  logger.log(level, `[RESPONSE] ${method} - ${success ? 'SUCCESS' : 'FAILED'} (${duration}ms)`, {
    success,
    duration,
    ...(process.env.LOG_LEVEL === 'debug' ? { data } : {}),
  });
};

export default logger;
