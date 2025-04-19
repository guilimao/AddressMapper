
# 将文字地址转换成地图上的坐标标记

## 用法：

1. **配置API密钥**  
   在 `tencent_api.txt` 中：
   - 第一行存入腾讯位置服务的 **API Key**  
   - 第二行存入 **Secret Key**  

   **示例**：

   JIRBZ-U2D3Z-JJPXW-7FHM4-BIPB6-ABCDE
   ketYpUGiZRIWFzr7Ow5W05GmXFABCDE

   > 可在 [腾讯位置服务控制台](https://lbs.qq.com/) 的 **控制台 → 应用管理 → 我的应用** 找到，Secret Key 需点击“编辑”查看。

2. **准备数据文件**  
   - 在 `1.xlsx` 中存入客户信息  
   - 格式需与 `Example.xlsx` 文件相同：  
     - 第一列：公司名称  
     - 第二列：公司地址  
     - 第三列：负责人姓名（可选，可留空）  

3. **运行程序**  
   双击运行 `AddressMapper.exe`，程序会自动生成 KML 文件。

4. **查看结果**  
   在 **Google Earth Pro** 中，通过菜单栏的 **文件 → 打开** 导入生成的 KML 文件。
