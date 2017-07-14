/*!
*   \file              ptkl_mf.h
*   
*  
*  \brief             Header for structures and defines of measurement feeder.
*  \author            Guan Zhen Chan \n
*  
*  \date              2017-07-03 Initial version \n
*  \version           0.01 , Initial version \n
*                     0.02 , Removed PARAM_CAPACITRANCE 0x71. PARAM_EXPECTED_RESULT is used instead.
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

    //Param: Get
        #define PARAM_SELFTEST_CONTACT_RES        0x80
        #define PARAM_SELFTEST_CAPACITY_CM_ID     0x81
        #define PARAM_SETUP_MEAS_COMP_TYPE        0x82
        #define PARAM_SETUP_MEAS_STRAY_CAPACITY   0x83
        #define PARAM_SETUP_MEAS_U                0x84
        #define PARAM_SETUP_MEAS_I                0x85
        #define PARAM_SETUP_MEAS_PHI              0x86
        #define PARAM_MEAS_COMPONENT_TYPE         0x87
        #define PARAM_MEAS_VALUE                  0x88
        #define PARAM_MEAS_U                      0x89
        #define PARAM_MEAS_I                      0x90
        #define PARAM_MEAS_PHI                    0x91
        #define PARAM_MEAS_FREQUENCY              0x92
        #define PARAM_MEAS_FWD_VOLTAGE            0x93
        #define PARAM_MEAS_FWD_CURRENT            0x94

    // Param: Defaults

        #define PARAM_DEFAULT_CURRENT             0x60
        #define PARAM_DEFAULT_VOTLAGE             0x61

    // Param : EEPROM Targets
        #define EEPROM_CM_1                       0x05
        #define EEPROM_CM_2                       0x06
        #define EEPROM_CM_3                       0x07

#endif /* PTKL_MF_H_ */
