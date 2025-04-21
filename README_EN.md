# Convert Text Addresses to Geographic Coordinates on a Map

## Usage:

1. **Create an API Key (This service is free)**  
   - Visit [Tencent Location Service](https://lbs.qq.com/), log in, and go to [Console → Application Management → My Applications](https://lbs.qq.com/dev/console/application/mine).  
   - Click "Add Key," enter a name (can be arbitrary), check "WebService API," click "Signature Verification," read and agree to the terms, then click "Add."  
   - After successfully creating the key, click ["Go to Allocate Quota"](https://lbs.qq.com/dev/console/quota/account), scroll to the second page, and click "Quota Allocation" under "Geocoding." Enable quota allocation for your newly created key and set an appropriate value (the default concurrency for this program is 5 requests per second), then submit.  
   - Go back to [Application Management → My Applications](https://lbs.qq.com/dev/console/application/mine), copy the **Key**, then click "Edit" and copy the **Secret Key**.  

2. **Configure the API Key**  
   - In `tencent_api.txt`:  
     - First line: Paste the **Key** from Tencent Location Service.  
     - Second line: Paste the **Secret Key**.  

   **Example**:  

   JIRBZ-U2D3Z-JJPXW-7FHM4-BIPB6-ABCDE  
   ketYpUGiZRIWFzr7Ow5W05GmXFABCDE  

3. **Prepare the Data File**  
   - Save address information in `1.xlsx`.  
   - The format should match `Example.xlsx`:  
     - First column: Company name  
     - Second column: Company address  
     - Third column: Contact person (optional, can be left blank)  

4. **Run the Program**  
   - Double-click `AddressMapper.exe`, and the program will automatically generate a KML file.  

5. **View the Results**  
   - Install [**Google Earth Pro**](https://support.google.com/earth/answer/168344#zippy=%2Cdownload-google-earth-pro-directly) and open the KML file with it.  
   - Alternatively, import the KML file in the [web version](https://earth.google.com/web) or [mobile app](https://play.google.com/store/apps/details?id=com.google.earth) of **Google Earth**.  
   - For users in China, [Ovi Maps](https://www.ovital.com/download/) is a better alternative.  

![Example Image](Effect.png)  
