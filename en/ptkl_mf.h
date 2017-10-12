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
*  \b Description: \n
*/

#ifndef PTKL_MF_H_
#define PTKL_MF_H_

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
    //Publics/Errors
        #define PUB_CM_NOT_CONNECTED              0xD1
        #define PUB_WRONG_POLARITY                0xD2
        #define PUB_MAX_VOLTAGE                   0xD3

    //Params : IOs

        #define PARAM_INP_COVER                   0x30

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
        #define PARAM_SELFTEST_CONTACT_RES        0x80
        #define PARAM_SELFTEST_CAPACITY           0x81
        #define PARAM_SELFTEST_RESISTANCE         0x82
        #define PARAM_SELFTEST_U                  0x83
        #define PARAM_SELFTEST_I                  0x84
        #define PARAM_SELFTEST_PHI                0x85
        #define PARAM_SELFTEST_FREQ               0x86
        #define PARAM_SETUP_MEAS_CONTACT_RES      0x87
        #define PARAM_SETUP_MEAS_CAPACITY         0x88
        #define PARAM_SETUP_MEAS_RESISTANCE       0x89
        #define PARAM_SETUP_MEAS_U                0x8A
        #define PARAM_SETUP_MEAS_I                0x8B
        #define PARAM_SETUP_MEAS_PHI              0x8C
        #define PARAM_SETUP_MEAS_FREQ             0x8D

        #define PARAM_MEAS_COMPONENT_TYPE1        0x90
        #define PARAM_MEAS_COMPONENT_TYPE2        0x91
        #define PARAM_MEAS_VALUE1                 0x92
        #define PARAM_MEAS_VALUE2                 0x93
        #define PARAM_MEAS_VALUE_MIN1             0x94
        #define PARAM_MEAS_VALUE_MAX1             0x95
        #define PARAM_MEAS_VALUE_MIN2             0x96
        #define PARAM_MEAS_VALUE_MAX2             0x97
        #define PARAM_MEAS_U                      0x98
        #define PARAM_MEAS_I                      0x99
        #define PARAM_MEAS_PHI                    0x9A
        #define PARAM_MEAS_FREQUENCY              0x9B
        #define PARAM_MEAS_FWD_VOLTAGE            0x9C
        #define PARAM_MEAS_FWD_CURRENT            0x9D

    // Param: Defaults

        #define PARAM_DEFAULT_CURRENT             0x60
        #define PARAM_DEFAULT_VOLTAGE             0x61
        #define MEAS_TOLERANCE_HIGHEND            0x62
        #define MEAS_TOLERANCE_MIDRANGE           0x63

    // Param : EEPROM Targets
        #define EEPROM_CM_1                       0x05
        #define EEPROM_CM_2                       0x06
        #define EEPROM_CM_3                       0x07

#endif
