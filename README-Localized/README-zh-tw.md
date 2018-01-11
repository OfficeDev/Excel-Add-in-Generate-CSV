# <a name="csv-generator-task-pane-add-in-sample-for-excel-2016"></a>Excel 2016 的 CSV 產生器工作窗格增益集範例

_適用於：Excel 2016_

這個工作窗格增益集示範如何使用 Excel 2016 中的 JavaScript API，從資料行名稱的清單建立資料表。共有兩種型態︰程式碼編輯器和 Visual Studio。

![CSV 產生器範例](../Images/ScreenCap1.PNG)

## <a name="try-it-out"></a>進行測試
### <a name="code-editor-version"></a>程式碼編輯器版本

部署及測試增益集的最簡單方式，是將檔案複製到網路共用。

1.  使用您所選擇的伺服器來裝載程式碼編輯器專案資料夾中的檔案。
2.  編輯資訊清單檔的 \<SourceLocation\> 和 \<Url\> 項目，讓它指向步驟 1 中建立的裝載位置。(例如，https://localhost/CSVGenerator/Home.html)
3.  將資訊清單 (TeacherCSVGenerator.xml) 複製到網路共用 (例如，\\\MyShare\MyManifests)。
4.  在 Excel 中，將包含資訊清單的共用位置新增為受信任的應用程式目錄。

    a.啟動 Excel，並開啟空白的試算表。

    b.選擇 **[檔案]** 索引標籤，然後選擇 **[選項]**。

    c.選擇 **[信任中心]**，然後選擇 **[信任中心設定]** 按鈕。

    d.選擇 **[受信任的增益集目錄]**。

    e.在 **[目錄 URL]** 方塊中，輸入您在步驟 3 建立的網路共用路徑，然後選擇 **[新增目錄]**。

   f.選取 **[顯示於功能表中]** 核取方塊，然後選擇 **[確定]**。接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。

5.  測試並執行增益集。

    a.在 Excel 2016 的 **[插入]** 索引標籤上，選擇 **[我的增益集]**。

    b.在 **[Office 增益集]** 對話方塊中，選擇 **[共用資料夾]**。

    c.選擇 **[教師 CSV 班級名冊範例]** > **[插入]**。增益集會在工作窗格中開啟，並在使用中工作表中建立 CSV 的班級名冊，如此螢幕擷取畫面所示。

   ![大學預算追蹤器範例](../Images/ScreenCap2.PNG)

    d.選擇 [教室管理服務]。

    e.按一下 [建立名冊] 按鈕，以在使用中工作表中插入空的名冊

      ![大學預算追蹤器範例](../Images/ScreenCap3.PNG)

    f.按一下 [Excel 匯出說明] 按鈕，瞭解如何將工作表匯出為 .csv 檔案。


### <a name="visual-studio-version"></a>Visual Studio 版本
1.  將專案複製到本機資料夾，並在 Visual Studio 中開啟 TeacherCSVGenerator.sln。
2.  按 F5 建置及部署範例增益集。Excel 會啟動，且增益集會在工作表右側空白部分的工作窗格中開啟，如下列螢幕擷取畫面所示。

  ![Excel CSV 產生器範例](../Images/ScreenCap1.PNG)

3.  從下拉式清單中選取線上教室管理服務
4.  藉由使用 **[產生名單]** 按鈕，新增學生名單資料表，然後查看使用中工作表中所建立的資料表。

  ![大學預算追蹤器範例](../Images/ScreenCap3.PNG)
5.  將學生新增至名單，方法是填入資料表標頭底下資料列中的儲存格。
6.  使用 Excel 中的匯出功能，將工作表儲存為 .csv 檔案。這個檔案的格式正確，可以匯入您所選擇的服務。


### <a name="learn-more"></a>Learn more

Excel JavaScript API 還有其他許多功能，可供您用於開發增益集。以下列出其中幾個可用的資源。

1.  [Excel 增益集程式設計概觀](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Excel 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel 增益集程式碼範例](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Excel 增益集 JavaScript API 參考](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [建立第一個 Excel 增益集](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)


此專案已採用 [Microsoft 開放原始碼管理辦法](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[管理辦法常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
