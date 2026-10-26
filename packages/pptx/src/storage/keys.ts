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
