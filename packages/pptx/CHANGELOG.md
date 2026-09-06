# @eflink-tech/pptx

## 0.2.0

### Minor Changes

- 4655e4b: 新增 `setPptxStorageBackend`：宿主可注入自定义存储后端对接远端 API（实现 put/get/remove/list 四个方法）；注入后 localStorage 崩溃镜像自动停用，仅保留当前文档 id 指针，远端为单一数据源。
