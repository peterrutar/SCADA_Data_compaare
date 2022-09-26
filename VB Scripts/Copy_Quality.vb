Sub Copy_Quality

HMIRuntime.Tags("Parameter_1").Write HMIRuntime.Tags("AI_DB_LIC_6_102_AI.PV").Read
HMIRuntime.Tags("Parameter_2").Write HMIRuntime.Tags("AI_DB_PIC_6_110_AI.PV").Read
HMIRuntime.Tags("Parameter_3").Write HMIRuntime.Tags("AI_DB_TT_6_211_AI.PV").Read
HMIRuntime.Tags("Parameter_4").Write HMIRuntime.Tags("AI_DB_TICR_6_302_AI.PV").Read
HMIRuntime.Tags("Parameter_5").Write HMIRuntime.Tags("AI_DB_QIC_6_402_T_AI.PV").Read
HMIRuntime.Tags("Parameter_6").Write HMIRuntime.Tags("AI_DB_QIC_6_403_T_AI.PV").Read
HMIRuntime.Tags("Parameter_7").Write HMIRuntime.Tags("AI_DB_TT_601_43_AI.PV").Read
HMIRuntime.Tags("Parameter_8").Write HMIRuntime.Tags("AI_DB_PIC_6_203_AI.PV").Read
HMIRuntime.Tags("Parameter_9").Write HMIRuntime.Tags("AI_DB_TIC_6_112_AI.PV").Read
HMIRuntime.Tags("Parameter_10").Write HMIRuntime.Tags("AI_DB_FIC_6_204_AI.PV").Read

End Sub