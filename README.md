## 环境配置
### 使用Conda一键安装
```bash
# 从environment.yml创建环境
conda env create -f environment.yml
# 激活环境
conda activate gui
# 安装补充依赖
pip install -r requirements.txt

```
## V2.0使用教程
### volume.xltm 配置
文件夹`volume.xltm`是Materialise Magics软件的输出报告模板文件，使用前需要将此文件放在一下目录，
```path
C:\ProgramData\Materialise\Magics\Templates\Materialise Magics\Office 2007-2013 Templates\Excel
```

### 体积信息导出
在Materialise Magics中，点击`分析&报告`->`生成报告`，在弹出的窗口中选择刚才放置的模板文件`volume.xltm`，点击`OK`，即可导出体积信息。

### 3dbudgcalc.exe使用
打开软件，点击`加载零件信息（xlsm）`，选择刚才导出的体积信息文件，填写MSC SliceViewer软件中计算的打印时间，点击`计算成本`，即可完成使用。