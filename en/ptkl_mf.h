/*!
*   \file              ptkl_mf.h
*   
*  
*  \brief             Header for structures and defines of measurement feeder.
*  \author            Guan Zhen Chan \n
*  
*  \date              2017-07-03 Initial version \n
*  \version           0.01 , Initial version \n
*                     2017-08-08
*                     0.02 , Removed PARAM_CAPACITRANCE 0x71. PARAM_EXPECTED_RESULT is used instead, enabled prepare commands.
*                     2017-09-29
*                     0.03 , Added PARAM_MEAS_VALUE_MIN and PARAM_MEAS_VALUE_MAX
*                     2017-10-05
*                     0.04 , Renumber params 0x80 - 0x97, to align with Interfacespec_MeasurementFeeder_MB_V3f.
*                     2017-11-14
*                     0.05 , Added default parameters 0x63 to 0x67
*                     2017-12-11
*                     0.06 , Added PARAM_DL_ZIEL_APPL_3 0x30, Calibration params 0x50 to 0x53, 
*                            Renumber PB_ERR params from 0xD? to 0x3?
*                     2017-12-27
*                     0.07 , Added PUB_ERROR_NOT_CAL, as an error for contact block not calibrated.
*                     2018-01-08
*                     0.08 , Added ACK_ERR_COVER_CLOSED
*                     2018-01-15
*                     v8 (0.09) , Reverted changes to be backward compatible with ptkl_mf.h V0.02.
*                            Newer params introduced in V0.03 toV0.08 are added on top of the changes.
*                     2018-02-22
*                     v9  , Added defines to access for new factory data: Tolerances / Offsets and settings.
*					                  Added define PARAM_MB_ERRORS
*					            2018-03-14
*					            v10 , Added  PARAM_MEAS_STATUS PARAM_MEAS_DEBUG
*  \b Description: \n
*/
    //Service Commands

    //Prepare Commands
        #define CMD_PREPARE_SETUP_MEASURE         0x41
        #define CMD_PREPARE_MEASURE               0x42
        #define CMD_PREPARE_MEASURE_AUTO          0x43
        #define CMD_PREPARE_SELFTEST              0x46
        #define CMD_PREPARE_CALIBRATION           0x47
    //Control Commands
    
    //Debug Commands

//PARAMS

    //Acknowledgements
        #define ACK_INVALID_CM                    0x50
        #define ACK_ERR_COVER_CLOSED              0x51
        #define ACK_ERR_SELFTEST_NOK              0x52
        
    //Publics/Errors
        #define PUB_CM_NOT_CONNECTED              0x31
        #define PUB_WRONG_POLARITY                0x32
        #define PUB_MAX_VOLTAGE                   0x33
        #define PUB_MB_ERROR                      0x34
        #define PUB_MB_CRC                        0x35
        #define PUB_ERROR_ST_CM_CAP               0x36
        #define PUB_ERROR_ST_CM_RES               0x37
        #define PUB_ERROR_ST_SC_RES               0x38
        #define PUB_ERROR_CM_CHANGED              0x39
        #define PUB_ERROR_MB_COMM                 0x3A

    //Params : IOs

        #define PARAM_INP_COVER                   0x30
        #define PARAM_MB_TEMP                     0x31
        
    // Params : Misc

        #define PARAM_MB_ERRORS                   0x54       
    // Param: Defaults

        #define PARAM_DEFAULT_CURRENT             0x60
        #define PARAM_DEFAULT_VOLTAGE             0x61
        #define PARAM_TOLERANCE_HIGHEND           0x62
        #define PARAM_TOLERANCE_MIDRANGE          0x63
        #define PARAM_DEFAULT_RES_ALL_WIRE        0x64
        #define PARAM_DEFAULT_RES_CONTACT_BLK     0x65
        #define PARAM_DEFAULT_CAP_CONTACT_BLK     0x66
        #define PARAM_DEFAULT_RES_SHORTCIRCUIT    0x67
        #define PARAM_DEFAULT_MEAS_TOLERANCE      0x68
        #define PARAM_DEFAULT_MEAS_OFFSET         0x69
        #define PARAM_DEFAULT_MEAS_SETTINGS       0x6A
        
    //Param: Send
        #define PARAM_MAX_VOLTAGE                 0x70
        #define PARAM_EXPECTED_RESULT             0x71
        #define PARAM_COMPONENT_TYPE              0x72
        #define PARAM_NUM_OF_CYCLES               0x73
        #define PARAM_POLARITY                    0x74
        #define PARAM_CURRENT                     0x75
        #define PARAM_VOLTAGE                     0x76
        #define PARAM_MODE_SELECT                 0x77 
        #define PARAM_NOMINAL_CAPACITY            0x78 

    //Param: Get
        #define PARAM_DL_ZIEL_APPL_3              0x30
       
        #define PARAM_SELFTEST_CONTACT_RES        0x80
        #define PARAM_SELFTEST_CAPACITY_CM_ID     0x81
        #define PARAM_SELFTEST_RESISTANCE         0x8C
        #define PARAM_SELFTEST_U                  0x8D
        #define PARAM_SELFTEST_I                  0x8E
        #define PARAM_SELFTEST_PHI                0x8F
        #define PARAM_SELFTEST_FREQ               0x95
        #define PARAM_SETUP_MEAS_COMP_TYPE        0x82
        #define PARAM_SETUP_MEAS_STRAY_CAPACITY   0x83
        #define PARAM_SETUP_MEAS_CONTACT_RES      0x96
        #define PARAM_SETUP_MEAS_RESISTANCE       0x97
        #define PARAM_SETUP_MEAS_U                0x84
        #define PARAM_SETUP_MEAS_I                0x85
        #define PARAM_SETUP_MEAS_PHI              0x86
        #define PARAM_SETUP_MEAS_FREQ             0x98

        #define PARAM_MEAS_COMPONENT_TYPE         0x87
        #define PARAM_MEAS_VALUE                  0x88
        #define PARAM_MEAS_VALUE_MIN              0x8A
        #define PARAM_MEAS_VALUE_MAX              0x8B
        #define PARAM_MEAS_COMPONENT_TYPE2        0x99
        #define PARAM_MEAS_VALUE2                 0x9A
        #define PARAM_MEAS_VALUE_MIN2             0x9B
        #define PARAM_MEAS_VALUE_MAX2             0x9C
        #define PARAM_MEAS_U                      0x89
        #define PARAM_MEAS_I                      0x90
        #define PARAM_MEAS_PHI                    0x91
        #define PARAM_MEAS_FREQUENCY              0x92
        #define PARAM_MEAS_FWD_VOLTAGE            0x93
        #define PARAM_MEAS_FWD_CURRENT            0x94
        
        #define PARAM_CAL_R_1KHz                  0x50
        #define PARAM_CAL_X_1KHz                  0x51
        #define PARAM_CAL_R_10KHz                 0x52
        #define PARAM_CAL_X_10KHz                 0x53

    // Param : EEPROM Targets
        #define EEPROM_CM_1                       0x05
        #define EEPROM_CM_2                       0x06
        #define EEPROM_CM_3                       0x07

    // Param: Debug parameters
        #define PARAM_MEAS_STATUS                 0x60
        #define PARAM_MEAS_DEBUG                  0x61
