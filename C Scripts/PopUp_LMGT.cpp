void OnAcquisitionClick(char* screenName, char* objectName)
{

	LINKINFO lplink;
      char *actorElement;	
	
	//Get the name of the current interface for the actor element and assign it to the TagPrefix
	////GetLink(screenName, objectName, "EnO_Name", &lplink);
	////SetPropChar(screenName,"EnS_ScreenWindow","TagPrefix",GetTagChar(lplink.szLinkName));	
	GetLink(screenName, objectName, "EnS_typeEnergyMeta", &lplink);
	actorElement = lplink.szLinkName;
	strcat(actorElement, ".name");
	SetPropChar(screenName,"EnS_ScreenWindow","TagPrefix",GetTagChar(actorElement));

	//Check the name of the assigned TagPrefix --> This command can be removed
	SetTagChar("TagPrefix", GetTagChar(actorElement));
	printf("TagName TagPrefix = %s\r\n", GetTagChar("TagPrefix"));
	//end

	SetPropChar(screenName,"EnS_ScreenWindow","ScreenName","EnS_EnergyDataBasic");
	SetPropBOOL(screenName,"EnS_ScreenWindow","Visible",TRUE);

}
