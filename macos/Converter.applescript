on run
    set itemsToConvert to choose file with prompt "选择要在本机转换的文件" with multiple selections allowed
    my convertFiles(itemsToConvert)
end run

on open itemsToConvert
    my convertFiles(itemsToConvert)
end open

on convertFiles(itemsToConvert)
    set modeKeys to {"pdf", "docx", "xlsx", "pptx", "docx-layout", "pptx-edit", "txt", "csv", "tsv", "odt", "ods", "odp", "rtf", "doc", "xls", "ppt", "png", "jpg", "webp", "bmp"}
    set modeNames to {"PDF 文档", "Word 文档 (.docx)", "Excel 工作簿 (.xlsx)", "PowerPoint 演示文稿 (.pptx)", "Word · 按页保留版式", "PowerPoint · 可编辑文字和表格", "纯文本 (.txt)", "逗号分隔表格 (.csv)", "制表符分隔表格 (.tsv)", "OpenDocument 文本 (.odt)", "OpenDocument 表格 (.ods)", "OpenDocument 演示 (.odp)", "富文本 (.rtf)", "Word 97–2003 (.doc)", "Excel 97–2003 (.xls)", "PowerPoint 97–2003 (.ppt)", "PNG 图片", "JPEG 图片", "WebP 图片", "BMP 图片"}
    set helperPath to POSIX path of (path to resource "convert.sh")
    set helperCommand to "/bin/sh " & quoted form of helperPath
    set argsText to ""
    repeat with sourceItem in itemsToConvert
        set argsText to argsText & " " & quoted form of (POSIX path of sourceItem)
    end repeat
    try
        set modesText to do shell script helperCommand & " --targets" & argsText
        if modesText is "" then
            display alert "这些文件没有共同的目标格式" message "请按文件类型分批选择。"
            return
        end if
        set possibleModes to paragraphs of modesText
        set namesToShow to {}
        repeat with i from 1 to count of modeKeys
            if (item i of modeKeys) is in possibleModes then set end of namesToShow to item i of modeNames
        end repeat
        set picked to choose from list namesToShow with title "本地文件转换器" with prompt "选择目标格式。结果保存在原文件旁。" default items {item 1 of namesToShow} OK button name "开始转换" cancel button name "取消"
        if picked is false then return
        repeat with i from 1 to count of modeNames
            if item i of modeNames is item 1 of picked then
                with timeout of 86400 seconds
                    set summaryText to do shell script helperCommand & " " & (item i of modeKeys) & argsText
                end timeout
                display dialog summaryText with title "本地文件转换器" buttons {"完成"} default button 1
                return
            end if
        end repeat
    on error errorText number errorNumber
        if errorNumber is not -128 then display alert "转换未完成" message errorText
    end try
end convertFiles
