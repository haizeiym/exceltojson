![表结构如下](./xlsx_format.png)

#### 第一行除缩进，导出文件头与导出目录为必须设置
#### 第二行开头为"#"的字段为不导出
#### s 导出到server目录
#### c 导出到client目录
#### sc 为全部导出
#### 以下值格式为pandas库中支持导出的格式

##### 使用方法：
`python exceltojson.py 文件路径`