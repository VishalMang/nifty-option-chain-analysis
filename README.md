# NIFTY Option Chain Analysis Template

Welcome to the NIFTY Option Chain Analysis Template repository! This project contains a custom MS Excel template designed to help you analyze real-time Nifty Option Chain data.

## Features

**• Real-Time Data Updates:** The template refreshes data every 3 minutes, providing the latest market insights.
 
**• In-depth Option Chain Analysis:** Explore open interest, volume, and implied volatility across different strike prices and expiry dates.

**• Trend Analysis:** Identify market sentiment and potential price movements using the latest trading data.

**• Option Strategy Comparison:** Compare popular strategies like straddles, strangles, and spreads to gauge market expectations and risk profiles.

**• Support & Resistance Levels:** Identify key support and resistance levels based on option chain data, aiding in short-term trading decisions.

## Dashboard
![15](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/c82862fa-3cfa-423c-84fc-c5f269f7c262)

![14](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/d0bc5115-54c2-4804-a87c-0b220bc55a7f)


## Important Note

While this template is designed to work efficiently, please be aware that it might be affected by cookies on different computers.

For detailed instructions on how to pull real-time data using cookies in Power Query within Excel, please refer to the provided screenshots and descriptions.

# Pulling Nifty Option Chain Data from NSE India to Excel
## Step-by-Step Guide
   **Step 1: Access the NSE India Website**
   1. Navigate to nseindia.com.
   2. On the homepage, select the "Derivatives" option.
   3. Click on any Nifty contract listed under all contracts.
      
![1_1](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/c9767a3f-8552-459b-b75a-b9efed74c612)




**Step 2: Inspect Network Traffic**
1. Right-click on the page and select "Inspect" to open the developer tools.
2. Go to the "Network" tab.

![2_2](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/62aeefce-847c-4b43-955e-c72eda5f11c9)


**Step 3: Capture the Cookie**
1. Click on the "Option Chain" link on the website.
2. In the "Network" tab, locate the network request for the option chain.
3. Click on this request, scroll down to the "Cookies" section, and copy the entire cookie content.

![3_2](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/7365ef7d-41e9-4d61-9f91-7102e45960cb)

**Step 4: Set Up Excel**
1. Open a blank Excel sheet.
2. In the first cell, type "cookie" and paste the copied cookie in the adjacent cell.
3. Select both cells, go to the "Data" tab, and choose "From Table/Range".
4. Ensure "My table has headers" is checked and click "OK".

![4](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/0d12dbb5-e4af-45a4-ba7a-95682db2c775)

**Step 5: Use Power Query Editor**
1. The Power Query Editor will open.
2. Drill down on the cookie column and rename it to "cookies".

![5](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/dd18e44d-61e0-4585-8dd2-2c1ce99e9626)

**Step 6: Modify the Query**
1. In the Power Query Editor, click on the "Home" button.
2. Open the "Advanced Editor".

![6](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/5917eb0c-2aa8-41b6-97a5-d6d9216564b5)

**Step 7: Edit the Syntax**
1. Copy the first and second lines of the syntax.
2. Paste these lines at the beginning and end of the syntax respectively.

![7](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/c64d126d-bf03-4b63-bc7f-d298d1d070cd)

**Step 8: Confirm Syntax**
1. Ensure no syntax errors are detected.

![8](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/e0e51d1e-d0aa-40ef-8b36-8a647efff6f6)

**Step 9: Create a New Query**
1. Right-click in the Queries pane and select "New Query".
2. Navigate to "Other Sources" and select "Blank Query".

![9](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/6b2b16cd-67cc-41aa-9ce6-7e546e5c35aa)

**Step 10: Add Syntax**
1. Copy the remaining syntax from the previous steps.
2. Paste it into the new blank query.

![10](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/7a88aff5-9755-4ec2-ae6d-cbd74aead73b)
![11](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/010f0aa6-6f90-441d-b48b-a825656e6afb)

**Step 11: Expand Columns**
1. The data will now display headers.
2. Expand the columns and select the required ones from both PE (Put) and CE (Call) sections.
3. Close and load the data.

![12](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/bb817920-4792-432c-af85-e13738bcf4df)

**Step 12: Utilize the Data**
1. The Excel sheet will display the Nifty option chain data.
2. You can now use this real-time data for analysis, including creating bar charts, graphs, and dashboards.

![13](https://github.com/VishalMang/nifty-option-chain-analysis/assets/164848822/e815e5b1-c161-4e43-9e6a-7ec8281b7f33)

## Power Query Text

let cookies=()=>

in 
    cookies


let
    Source = Json.Document(Web.Contents("https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY", [Headers=[#"Accept-Encoding"="gzip, deflate", #"Accept-Language"="en-US,en", #"User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36 OPR/68.0.3618.206", Cookie="8A87B46ABA4CAFF3B69F913708597828~xM8YKIspOw4k2OTekhqVI8Ft8AHi/RYKbvLqfKirkbafd1XOqJELenPBKr4Y+FAgbqei34v6NKmWyp1RWvhDhh2jrLXAela7ZdmyrHShEPCaVopVDul8R91B2SbFshwrUsS7yKn5+cmpmaF25zGeiAjHTfTMnii7F2E1slCnkEo="]])),
    records = Source[records],
    data = records[data],
    #"Converted to Table" = Table.FromList(data, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"strikePrice", "expiryDate", "CE", "PE"}, {"strikePrice", "expiryDate", "CE", "PE"})
in
    #"Expanded Column1"






## Disclaimer
This template is provided as-is, and the creator is not responsible for any issues arising from its use. Use with caution and verify the data independently.

## Contact
Feel free to reach out if you have any questions or need further assistance:

• Vishal Mang

• LinkedIn: linkedin.com/in/vishal-mang-a7983a154
