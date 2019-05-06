UT_AutoReport
---

A (trying to be) convient report generator in UTECHZONE

Workflow --
    1. Write your report on pure text format,
    2. UT_AutoReport will generate the report for you.
    Done!

Supported Format --
    Currently support .json file.

    Usage of .json:
        Basically your report consist of two kinds of text information,
        1. define a Project by declaring its name, which is the project/item/tasks you are working on
        2. describe one or more Progress, and relate Progress to one defined Project

        Project is basically just a name.
        Progress has many fields to fill out, like date, decription text, pictures, ...
        with most of them are only optional.

        UT_AutoReport can combine multiple .json files,
        Project has to be decalred in at least one of the files,
        Progress can scatter over the files, so one can write .json with ease

    Please checkout template/ for detailed example

Supported reports --
    Weekly report:
        # Serve as both statement and summary of the progress of the week
        # Format:
            a. One .xlsx file for summary of the progress
            b. One .pptx file for detailed report on the progress