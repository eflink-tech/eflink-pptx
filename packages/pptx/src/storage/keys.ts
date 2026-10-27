/** DocuLuma 命名空间的 localStorage 键（新键 + 迁移输入用的旧键），集中在此避免行内字面量散落。 */

/** 当前文档指针（上次打开的文档 id） */
export const LAST_DOC_KEY = 'doculuma:pptx:last-doc'
/** 崩溃恢复用的本地草稿镜像 */
export const MIRROR_KEY = 'doculuma:pptx:mirror'

/** 迁移输入：DocuLuma 之前的 efLink 品牌键名（只读，迁移后删除） */
export const LEGACY_LAST_DOC_KEY = 'eflink-pptx-last-doc'
export const LEGACY_MIRROR_KEY = 'eflink-pptx-mirror'

/** 迁移输入：旧 IndexedDB 库名（迁移后保留，不删除） */
export const LEGACY_FILES_DB_NAME = 'eflink-pptx-files'
export const LEGACY_AI_DB_NAME = 'eflink-pptx-ai'

/** 旧库一次性导入标记（按表落键）：导入成功或确认跳过后置位，防"删光数据后旧数据复活" */
export const FILES_IMPORT_DONE_KEY = 'doculuma:pptx:legacy-imported:documents'
export const AI_CONVERSATIONS_IMPORT_DONE_KEY = 'doculuma:pptx:legacy-ai-imported:conversations'
export const AI_SETTINGS_IMPORT_DONE_KEY = 'doculuma:pptx:legacy-ai-imported:settings'
export const AI_TEMPLATES_IMPORT_DONE_KEY = 'doculuma:pptx:legacy-ai-imported:templates'
