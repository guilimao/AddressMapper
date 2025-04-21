
# 将文字地址转换成地图上的坐标标记

## 用法：

1. **创建API密钥（该服务是免费的）**
   - 打开[腾讯位置服务](https://lbs.qq.com/) ，登录，点击[控制台 → 应用管理 → 我的应用](https://lbs.qq.com/dev/console/application/mine)，  
   - 点击添加 key，输入名称（可随意填写），勾选 “WebService API” ，点击“签名校验”，“阅读并同意”，最后点“添加”即可；   
   - 创建 key 成功，点击[“立即前往分配”](https://lbs.qq.com/dev/console/quota/account)，翻到第二页，点击“地址解析”的“配额分配”，打开你刚刚创建的 key 的配额分配，并填入一个比较合适的值（本程序默认的并发量是一秒5次），提交。
   - 打开[应用管理 → 我的应用](https://lbs.qq.com/dev/console/application/mine)，把所需的 **key** 复制下来，再点击编辑，把 **Secret Key** 复制下来。
   

2. **配置API密钥**  
   - 在 `tencent_api.txt` 中：
     - 第一行存入腾讯位置服务的 **Key**  
     - 第二行存入 **Secret Key**  

   **示例**：

   JIRBZ-U2D3Z-JJPXW-7FHM4-BIPB6-ABCDE   
   ketYpUGiZRIWFzr7Ow5W05GmXFABCDE

4. **准备数据文件**  
   - 在 `1.xlsx` 中存入地址信息  
   - 格式需与 `Example.xlsx` 文件相同：  
     - 第一列：公司名称  
     - 第二列：公司地址  
     - 第三列：负责人姓名（可选，可留空）  

5. **运行程序**  
   - 双击运行 `AddressMapper.exe`，程序会自动生成 KML 文件。

6. **查看结果**  
   - 安装[**Google Earth Pro**](https://support.google.com/earth/answer/168344#zippy=%2C%E4%B8%8B%E8%BD%BD-google-%E5%9C%B0%E7%90%83-pro-%E7%9B%B4%E6%8E%A5%E5%AE%89%E8%A3%85%E7%A8%8B%E5%BA%8F)，双击 KML 文件，用**Google Earth Pro**打开。
   - 也可以选择在[网页端](https://earth.google.com/web)/[移动端](https://play.google.com/store/apps/details?id=com.google.earth)的 **Google Earth** 中导入 KML 文件。
   - 对于中国用户，[奥维互动地图](https://www.ovital.com/download/) 是个更好的选择。

![效果图片](Effect.png)
