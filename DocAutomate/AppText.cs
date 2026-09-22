using System.Collections.Generic;
using System.Globalization;

namespace DocAutomate
{
    // Application-owned text; OS dialogs and exception details follow Windows.
    internal static class AppText
    {
        internal static bool IsKorean = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "ko";
        private static readonly Dictionary<string, string> Korean = new Dictionary<string, string>
        {
            { "Excel Document Generator", "Excel \ubb38\uc11c \uc0dd\uc131\uae30" },
            { "Browse", "\ucc3e\uc544\ubcf4\uae30" },
            { "Generate", "\uc0dd\uc131" },
            { "Excel file", "Excel \ud30c\uc77c" },
            { "PowerPoint file", "PowerPoint \ud30c\uc77c" },
            { "Model", "\ubaa8\ub378" },
            { "Part", "\ubd80\ud488 \ubc88\ud638" },
            { "CRC", "CRC" },
            { "Select an Excel file", "Excel \ud30c\uc77c \uc120\ud0dd" },
            { "Excel workbooks (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls", "Excel \ud1b5\ud569 \ubb38\uc11c (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls" },
            { "Select the Software changes PowerPoint", "\uc18c\ud504\ud2b8\uc6e8\uc5b4 \ubcc0\uacbd \uc0ac\ud56d PowerPoint \uc120\ud0dd" },
            { "PowerPoint presentations (*.pptx)|*.pptx", "PowerPoint \ud504\ub808\uc820\ud14c\uc774\uc158 (*.pptx)|*.pptx" },
            { "Please select an Excel file first.", "\uba3c\uc800 Excel \ud30c\uc77c\uc744 \uc120\ud0dd\ud558\uc138\uc694." },
            { "Select a file", "\ud30c\uc77c \uc120\ud0dd" },
            { "Please enter a Model that can be used in a file name (no \\/:*?\"<>| characters or trailing period).", "\ud30c\uc77c \uc774\ub984\uc5d0 \uc0ac\uc6a9\ud560 \uc218 \uc788\ub294 \ubaa8\ub378\uba85\uc744 \uc785\ub825\ud558\uc138\uc694 (\\/:*?\"<>| \ubb38\uc790 \ubc0f \ub05d\uc758 \ub9c8\uce68\ud45c\ub294 \uc0ac\uc6a9\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4)." },
            { "Invalid Model", "\uc798\ubabb\ub41c \ubaa8\ub378\uba85" },
            { "Please enter a valid name in {0}.", "{0}\uc5d0 \uc62c\ubc14\ub978 \uc774\ub984\uc744 \uc785\ub825\ud558\uc138\uc694." },
            { "Invalid file name", "\uc798\ubabb\ub41c \ud30c\uc77c \uc774\ub984" },
            { "{0} must start with {1}.", "{0}\uc740(\ub294) {1}(\uc73c)\ub85c \uc2dc\uc791\ud574\uc57c \ud569\ub2c8\ub2e4." },
            { "Invalid {0}", "\uc798\ubabb\ub41c {0}" },
            { "Please select today or a future date.", "\uc624\ub298 \ub610\ub294 \uc774\ud6c4 \ub0a0\uc9dc\ub97c \uc120\ud0dd\ud558\uc138\uc694." },
            { "Invalid date", "\uc798\ubabb\ub41c \ub0a0\uc9dc" },
            { "A file or folder already exists at:\n{0}\nPlease enter a different name.", "\ub2e4\uc74c \uc704\uce58\uc5d0 \ud30c\uc77c \ub610\ub294 \ud3f4\ub354\uac00 \uc774\ubbf8 \uc788\uc2b5\ub2c8\ub2e4:\n{0}\n\ub2e4\ub978 \uc774\ub984\uc744 \uc785\ub825\ud558\uc138\uc694." },
            { "Name already exists", "\uc774\ubbf8 \uc874\uc7ac\ud558\ub294 \uc774\ub984" },
            { "Excel document generated:\n{0}\n\nCould not open its folder.\n{1}", "Excel \ubb38\uc11c\uac00 \uc0dd\uc131\ub418\uc5c8\uc2b5\ub2c8\ub2e4:\n{0}\n\n\ud3f4\ub354\ub97c \uc5f4 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4.\n{1}" },
            { "Document generated", "\ubb38\uc11c \uc0dd\uc131 \uc644\ub8cc" },
            { "Could not generate the Excel document.\n{0}", "Excel \ubb38\uc11c\ub97c \uc0dd\uc131\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4.\n{0}" },
            { "Generate failed", "\uc0dd\uc131 \uc2e4\ud328" },
            { "The generated file must use the same Excel format as the source.", "\uc0dd\uc131\ud560 \ud30c\uc77c\uc740 \uc6d0\ubcf8\uacfc \ub3d9\uc77c\ud55c Excel \ud615\uc2dd\uc774\uc5b4\uc57c \ud569\ub2c8\ub2e4." },
            { "Workbook updates require .xlsx or .xlsm. Save the .xls workbook in a modern Excel format first.", "\ud1b5\ud569 \ubb38\uc11c\ub97c \uc5c5\ub370\uc774\ud2b8\ud558\ub824\uba74 .xlsx \ub610\ub294 .xlsm \ud615\uc2dd\uc774 \ud544\uc694\ud569\ub2c8\ub2e4. \uba3c\uc800 .xls \ud30c\uc77c\uc744 \ud574\ub2f9 \ud615\uc2dd\uc73c\ub85c \uc800\uc7a5\ud558\uc138\uc694." },
            { "Please select an .xlsx, .xlsm, or .xls file.", ".xlsx, .xlsm \ub610\ub294 .xls \ud30c\uc77c\uc744 \uc120\ud0dd\ud558\uc138\uc694." },
            { "Ambiguous group or setting headers.", "\uadf8\ub8f9 \ub610\ub294 \uc124\uc815 \uba38\ub9ac\uae00\uc744 \uba85\ud655\ud558\uac8c \uc2dd\ubcc4\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { "Please select a .pptx PowerPoint file.", ".pptx PowerPoint \ud30c\uc77c\uc744 \uc120\ud0dd\ud558\uc138\uc694." },
            { ", table ", ", \ud45c " },
            { ", row ", ", \ud589 " },
            { ": ambiguous value headers.", ": \uac12 \uba38\ub9ac\uae00\uc744 \uba85\ud655\ud558\uac8c \uc2dd\ubcc4\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { ": cannot identify As Is and To Be columns.", ": \ubcc0\uacbd \uc804 \ubc0f \ubcc0\uacbd \ud6c4 \uc5f4\uc744 \uc2dd\ubcc4\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { ": group and setting columns require explicit headers.", ": \uadf8\ub8f9 \ubc0f \uc124\uc815 \uc5f4\uc5d0 \uba85\ud655\ud55c \uba38\ub9ac\uae00\uc774 \ud544\uc694\ud569\ub2c8\ub2e4." },
            { ": missing required columns.", ": \ud544\uc218 \uc5f4\uc774 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { ": a blank group must be vertically merged.", ": \ube48 \uadf8\ub8f9 \uc140\uc740 \uc138\ub85c\ub85c \ubcd1\ud569\ub418\uc5b4 \uc788\uc5b4\uc57c \ud569\ub2c8\ub2e4." },
            { ": missing group, setting, As Is or To Be value.", ": \uadf8\ub8f9, \uc124\uc815, \ubcc0\uacbd \uc804 \ub610\ub294 \ubcc0\uacbd \ud6c4 \uac12\uc774 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { ": duplicate setting ", ": \uc911\ubcf5\ub41c \uc124\uc815 " },
            { "No data table found on a Software Changes slide.", "\uc18c\ud504\ud2b8\uc6e8\uc5b4 \ubcc0\uacbd \uc0ac\ud56d \uc2ac\ub77c\uc774\ub4dc\uc5d0\uc11c \ub370\uc774\ud130 \ud45c\ub97c \ucc3e\uc744 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { "The workbook is missing required cell formatting definitions.", "\ud1b5\ud569 \ubb38\uc11c\uc5d0 \ud544\uc218 \uc140 \uc11c\uc2dd \uc815\uc758\uac00 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { "Expected one Excel match for {0}; found {1}.", "{0}\uc5d0 \ud574\ub2f9\ud558\ub294 Excel \uc140\uc774 \ud558\ub098\uc5ec\uc57c \ud558\uc9c0\ub9cc {1}\uac1c\ub97c \ucc3e\uc558\uc2b5\ub2c8\ub2e4." },
            { "Cannot replace a formula for {0}.", "{0}\uc758 \uc218\uc2dd\uc744 \ubc14\uafc0 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4." },
            { "Excel value for {0} does not match PowerPoint As Is ({1}).", "{0}\uc758 Excel \uac12\uc774 PowerPoint\uc758 \ubcc0\uacbd \uc804 \uac12({1})\uacfc \uc77c\uce58\ud558\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4." },
            { "Multiple changes target the same Excel cell: {0}", "\uc5ec\ub7ec \ubcc0\uacbd \uc0ac\ud56d\uc774 \ub3d9\uc77c\ud55c Excel \uc140\uc744 \ub300\uc0c1\uc73c\ub85c \ud569\ub2c8\ub2e4: {0}" }
        };

        internal static string Get(string english, params object[] arguments)
        {
            string translated;
            string text = IsKorean && Korean.TryGetValue(english, out translated) ? translated : english;
            return arguments.Length == 0 ? text : string.Format(CultureInfo.CurrentCulture, text, arguments);
        }
    }
}
