Summoner's War 按鍵精靈 (VB.Net)
==================

**NOTE:** Demonstrate how this program works and what we should do in advance.
================

*   [開發環境](#environment)

*   [程式介紹](#introduce)

*   [Step](#step)

*   [Key Action](#action)

* * *


<h1 id="environment">開發環境</h1>

Window 7 + Visual Studio (.net 4.5.2) + Bluestacks。

<h1 id="introduce">程式介紹</h1>

魔靈召喚自動刷地城、火山腳本輔助程式

依 Bluestack 擷圖判斷作為判斷條件，對 Bluestack 模擬器內發送鍵盤操作訊息，完成自動化腳本輔助程式。

詳情參照 [Medium]
  [Medium]: https://medium.com/@replacedtoy/%E7%94%A8-gdi-%E5%B0%B1%E5%AF%AB%E5%87%BA%E8%87%AA%E5%8B%95%E5%8C%96%E9%81%8A%E6%88%B2%E8%85%B3%E6%9C%AC-ft-bluestack-10fc8a6da4d2

<h1 id="step">Step</h1>

切記模擬器選項：繪圖方式要用 DirectX，切勿使用 OpenGL。

開啟模擬器、最大化

開啟自動化腳本輔助程式

<h1 id="action">Key Action : </h1>
F1 : 按鍵判圖點擊開啟。
F2 : 按鍵判圖點擊關閉。
F3 : 擷取當前畫面並與第 Index 張影像比對 
F4 : 擷取當前畫面並存檔
F5 : Index += 1 
F6 : Index = 0
F7 : 模擬當前 Index 所對應點擊 Bluestack 的事件