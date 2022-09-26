#include "apdefap.h"
void OnClick(char* lpszPictureName, char* lpszObjectName, char* lpszPropertyName)
{
LINKINFO fplink;
int i;
int j;
char* Prefix = SysMalloc(20);
char* Prefix2 = SysMalloc(20);
char* Suffix = SysMalloc(10);
char* Suffix2 = SysMalloc(10);

GetLink(lpszPictureName, lpszObjectName,"AI" , &fplink);    //"" -> Tag type

strcpy (Prefix, "PopUp_Settings_"); 

for (i = 1; i < 6; ++i){
    switch(i){
        case 1:
            strcpy (Suffix, "1");
                if (!(GetPropBOOL(lpszPictureName,Prefix,"Visible"))){
                    //Check if PopUp for this Faceplate is allready open elsewhere
                    for (j = (i-1); j > 0; --j){
                        switch(j){
                            case 1:
                                strcpy (Suffix2, "1");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 2:
                                strcpy (Suffix2, "2");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 3:
                                strcpy (Suffix2, "3");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 4:
                                strcpy (Suffix2, "4");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 5:
                                strcpy (Suffix2, "5");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                        }
                    }
                    SetPropChar(lpszPictureName,Prefix,"TagPrefix",fplink.szLinkName);
                    SetPictureName(lpszPictureName,Prefix,"AI_PopUp\\AI_Settings.PDL");
                    SetPropBOOL(lpszPictureName,Prefix,"Visible",TRUE);
                    SetPropChar(lpszPictureName,"txt","Text",Prefix);
                    exit(0);
                }
        case 2:
            strcpy (Suffix, "2");
            if (!(GetPropBOOL(lpszPictureName,Prefix,"Visible"))){
                    //Check if PopUp for this Faceplate is allready open elsewhere
                    for (j = (i-1); j > 0; --j){
                        switch(j){
                            case 1:
                                strcpy (Suffix2, "1");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 2:
                                strcpy (Suffix2, "2");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 3:
                                strcpy (Suffix2, "3");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 4:
                                strcpy (Suffix2, "4");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 5:
                                strcpy (Suffix2, "5");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                        }
                    }
                    SetPropChar(lpszPictureName,Prefix,"TagPrefix",fplink.szLinkName);
                    SetPictureName(lpszPictureName,Prefix,"AI_PopUp\\AI_Settings.PDL");
                    SetPropBOOL(lpszPictureName,Prefix,"Visible",TRUE);
                    SetPropChar(lpszPictureName,"txt","Text",Prefix);
                    exit(0);
                }
        case 3:
            strcpy (Suffix, "3");
            if (!(GetPropBOOL(lpszPictureName,Prefix,"Visible"))){
                    //Check if PopUp for this Faceplate is allready open elsewhere
                    for (j = (i-1); j > 0; --j){
                        switch(j){
                            case 1:
                                strcpy (Suffix2, "1");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 2:
                                strcpy (Suffix2, "2");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 3:
                                strcpy (Suffix2, "3");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 4:
                                strcpy (Suffix2, "4");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 5:
                                strcpy (Suffix2, "5");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                        }
                    }
                    SetPropChar(lpszPictureName,Prefix,"TagPrefix",fplink.szLinkName);
                    SetPictureName(lpszPictureName,Prefix,"AI_PopUp\\AI_Settings.PDL");
                    SetPropBOOL(lpszPictureName,Prefix,"Visible",TRUE);
                    SetPropChar(lpszPictureName,"txt","Text",Prefix);
                    exit(0);
                }
        case 4:
            strcpy (Suffix, "4");
            if (!(GetPropBOOL(lpszPictureName,Prefix,"Visible"))){
                    //Check if PopUp for this Faceplate is allready open elsewhere
                    for (j = (i-1); j > 0; --j){
                        switch(j){
                            case 1:
                                strcpy (Suffix2, "1");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 2:
                                strcpy (Suffix2, "2");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 3:
                                strcpy (Suffix2, "3");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 4:
                                strcpy (Suffix2, "4");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 5:
                                strcpy (Suffix2, "5");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                        }
                    }
                    SetPropChar(lpszPictureName,Prefix,"TagPrefix",fplink.szLinkName);
                    SetPictureName(lpszPictureName,Prefix,"AI_PopUp\\AI_Settings.PDL");
                    SetPropBOOL(lpszPictureName,Prefix,"Visible",TRUE);
                    SetPropChar(lpszPictureName,"txt","Text",Prefix);
                    exit(0);
                }
        case 5:
            strcpy (Suffix, "5");
            if (!(GetPropBOOL(lpszPictureName,Prefix,"Visible"))){
                    //Check if PopUp for this Faceplate is allready open elsewhere
                    for (j = (i-1); j > 0; --j){
                        switch(j){
                            case 1:
                                strcpy (Suffix2, "1");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 2:
                                strcpy (Suffix2, "2");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 3:
                                strcpy (Suffix2, "3");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 4:
                                strcpy (Suffix2, "4");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                            case 5:
                                strcpy (Suffix2, "5");
                                strncat(Prefix2, Suffix2, 1);
                                if ((GetPropChar(lpszPictureName,Prefix2,"TagPrefix")) = fplink.szLinkName){
                                    exit(0);
                                }
                        }
                    }
                    SetPropChar(lpszPictureName,Prefix,"TagPrefix",fplink.szLinkName);
                    SetPictureName(lpszPictureName,Prefix,"AI_PopUp\\AI_Settings.PDL");
                    SetPropBOOL(lpszPictureName,Prefix,"Visible",TRUE);
                    SetPropChar(lpszPictureName,"txt","Text",Prefix);
                    exit(0);
                }
    }

    // Combine the PopUp name
    strncat(Prefix, Suffix, 1);

    
    }
}
}