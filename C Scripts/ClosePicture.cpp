#include "apdefap.h"
void OnClick(char* lpszPictureName, char* lpszObjectName, char* lpszPropertyName)
{
    SetPropBOOL(GetParentPicture(lpszPictureName),"Alarms","Visible",FALSE);
} 