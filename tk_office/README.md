# pyinstaller 资源包
1. pyinstaller -F xxx.py  初步打包。然后将 build、dist目录删除
2. 修改生成的xxx.spec, 将 Analysis 下 datas 配置修改。例子datas=[('res/conf.ini', 'res'),('res/tempdoc.docx', 'res'),('res/log_err.txt', 'res')],
是将 项目中的res目录下conf.ini 初步打包到打包项目中的res目录下。
3. 重新执行pyinstaller -F xxx.spec

