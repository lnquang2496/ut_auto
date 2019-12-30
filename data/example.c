void test_{func}(){
    struct CPPTH_LOOP_INPUT_STRUCT {
        /* Test case data declarations */
        char* name;
        char* description;
        char* expected_calls;
        int execute;
        {data_declarations}
        int32_t var1;
        uint32_t var2;
        uint32_t * var3;
        int32_t var4;
        int32_t * var5;
        struct def * var_def;
        int32_t var6;
        int32_t * var7;
        int32_t var8;
        int32_t * var9;
        {return_data_declarations}
        int32_t expected_ReturnValue;
    };
    {expected_return_data_declarations}
    int32_t ReturnValue;
    /* Import external data declarations */
    #include "test_{func}.h"
    {external_data_declarations}
    uint32_t var3;
    struct abc var_abc;
    int32_t var5;
    struct def var_def;
    int32_t var7;
    int32_t var9;
    START_TEST_LOOP();
        /* Expected Call Sequence  */
        EXPECTED_CALLS(CURRENT_TEST.expected_calls);
            {external_data_initialize}
            if (CURRENT_TEST.var3 != NULL) {
                CURRENT_TEST.var3 = &var3;
            }

            var_abc.var4 = CURRENT_TEST.var4;

            if (CURRENT_TEST.var5 != NULL) {
                CURRENT_TEST.var5 = &var5;
            }
            var_abc.var5 = CURRENT_TEST.var5;

            if (CURRENT_TEST.var_def != NULL) {
                CURRENT_TEST.var_def = &var_def;
            }
            var_abc.var_def = CURRENT_TEST.var_def;

            var_def.var6 = CURRENT_TEST.var6;

            if (CURRENT_TEST.var7 != NULL) {
                CURRENT_TEST.var7 = &var7;
            }
            var_def.var7 = CURRENT_TEST.var7;

            var_abc.var_ghi.var8 = CURRENT_TEST.var8;

            if (CURRENT_TEST.var9 != NULL) {
                CURRENT_TEST.var9 = &var9;
            }
            var_abc.var_ghi.var9 = CURRENT_TEST.var9;
            /* Call SUT */
            {expected_return_assign}{func}({argument});
            ReturnValue = {func}(var1, var2, var3, var_abc);
            /* Test case checks */
            {expected_data_check}

            {expected_return_check}
            CHECK_S_INT(ReturnValue , expected_ReturnValue);
        END_CALLS();
    END_TEST_LOOP();
}
