/**
 * Input: 通用工具参数中的 app（word/excel/ppt）
 * Output: 带 targetApp 的 WPS 调用参数
 * Pos: 通用工具目标组件参数。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 通用工具的“目标组件”参数。
 * 保存、另存、导出 PDF 等命令三个组件都能执行。Mac 版同时开着几种文件时，
 * 用 app 指定对哪个组件操作；不填就发给开着文件、最近操作过的那个组件。
 */
import { ToolParameterSchema } from '../../types/tools';

export const APP_TARGET_PROPERTY: Record<string, ToolParameterSchema> = {
  app: {
    type: 'string',
    enum: ['word', 'excel', 'ppt'],
    description: '对哪个组件操作：word=文字、excel=表格、ppt=演示。不填则用开着文件、最近操作过的那个',
  },
};

/** 把工具参数里的 app 转成发给 WPS 的 targetApp */
export function withTarget(params: Record<string, unknown>, args: Record<string, unknown>): Record<string, unknown> {
  const app = typeof args.app === 'string' ? args.app : '';
  return app ? { ...params, targetApp: app } : params;
}
