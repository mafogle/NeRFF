# NeRFF
New Residential Fiber Intake-Form NTInet

The purpose of this program is to create a customer intake form that has EULA/ToS, pricing, and account information while simulatneously on a sales call. 

Features that still need to be added:
  1.) Reliable docx to pdf conversion
  2.) Scheduled installation date
  3.) First month Pro Rate calculator that automatically calculates by finding integer value of last day int(%d) in scheduledInstall month and dividing monthly rate(s)         
      by that amount then multiplying the quotient by (int(monthLastDay) - int(installDateDay))
  4.) Establish Total Due on Installation by prorating servicePlan, voipService, smartWifi, adding additionalItems, salesTax, and installationFee.
      
      totalInstallAmount = (serviceRate + voipService + smartWifi + wifiExtender) / (monthLastDay - installDateDay) + ((additionalItem1 + additionaitem2 + additionalItem3) * (1 + salesTax)) + installationFee
      
  5.) Add fields to Property details to answer Custom Fields in SF
  6.) Open, print, and overwrite Notepad with info to be copy and pasted into Plat, Plume, SmartOLT, Service Fusion.
