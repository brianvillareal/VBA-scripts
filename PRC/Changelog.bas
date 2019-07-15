Attribute VB_Name = "Changelog"
'* Brian Villareal.
'* Changelog.
'*  1.01 - 01/31/2019:
'*      Scientific notation will be removed when concatenating columns B and A (PRC, ConvertPRCtoWaystar).
'*      AutoFilter criteria will grab everything except "Child" (PRC, DeleteParent).
'*  1.02 - 01/31/2019:
'*      strTemplate updated for new template file (PRC, WaystarJE).
'*      Column identifiers updated for wbkWaystar to match new template file (PRC, Waystar JE).
'*      Auto_Add Subroutine included to prevent error with automatic install (CallbackPRC).
'*  1.03 - 02/01/2019:
'*      Worksheet references added for new template file (PRC, WaystarJE).
'*      Source data will be copied to wbkSource (PRC, CovertPRCtoWaystar).
'*      wbkSource will be used in each JE file (PRC, WaystarJE).
'*      wbkSource will be closed prior to ending the parent subroutine (PRC, ConverPRCtoWaystar).
'*      Journal Date will populate on first tab of template file (PRC, WaystarJE).
'*  1.04 - 02/07/2019:
'*      Data pulled will depend on user date input (PRC, ConvertPRCtoWaystar).
'*      DeleteUserExclusion subroutine added to remove unnecessary data (PRC, ConvertPRCtoWaystar).
'*      If statement causing credits to flip incorrectly was removed (PRC, ConvertPRCtoWaystar).
'*      Use of lngFirstGL and lngSecondGL was switched (PRC, ConvertPRCtoWaystar).
'*  1.05 - 03/11/2019:
'*      Fully qualified all workbook references (PRC).
'*      Split module into separate subroutines (PRC).
'*      User will be alerted and subroutine will end if there is nothing to post (CheckRedundancy).
'*      Added functionality for posting a range of dates (GetStartDate, GetEndDate, ParseDateRange).
'*      Unmapped activity will be deleted instead of causing the macro to prematurely end (DeletePaymentAccounts, MapBanks, MapPayments).
'*  1.06 - 03/18/2019:
'*      "PLB" added as Case (MapPayments).
'*  1.07 - 04/24/2019:
'*      Added PRC15Characters module.
