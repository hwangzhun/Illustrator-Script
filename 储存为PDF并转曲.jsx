/**
 * @fileoverview Illustrator 一键转曲将链接图嵌入文档导出pdf文件
 * 
 * @author Hwangzhun <huangzhenmsn@hotmail.com>
 * @version v0.3
 * @date 2025/4/24
 * 
 * @description 将当前文档的链接图嵌入到文档，并将文字转曲导出pdf文件
 */

// 自定义提示函数
function showAlert(message, title) {
    var win = new Window('dialog', title || '提示');
    win.add('statictext', undefined, message);
    var btn = win.add('button', undefined, '确定');
    btn.onClick = function() { win.close(); };
    win.show();
}

// 自定义确认函数
function showConfirm(message, title) {
    var win = new Window('dialog', title || '确认');
    win.add('statictext', undefined, message);
    var btnGroup = win.add('group');
    var yesBtn = btnGroup.add('button', undefined, '是');
    var noBtn = btnGroup.add('button', undefined, '否');
    
    var result = false;
    yesBtn.onClick = function() { result = true; win.close(); };
    noBtn.onClick = function() { result = false; win.close(); };
    
    win.show();
    return result;
}

// 显示设置对话框
var dialog = new Window('dialog', '导出PDF设置');
dialog.orientation = 'column';
dialog.alignChildren = ['left', 'top'];

// 导出位置
var locationGroup = dialog.add('group');
locationGroup.add('statictext', undefined, '导出位置:');
var locationDropdown = locationGroup.add('dropdownlist', undefined, ['同目录下', '选择其他位置']);
locationDropdown.selection = 0;

// 导出后缀
var suffixGroup = dialog.add('group');
suffixGroup.add('statictext', undefined, '导出后缀:');
var suffixInput = suffixGroup.add('edittext', undefined, '-转曲');
suffixInput.characters = 10;

// 转曲选项
var outlineCheck = dialog.add('checkbox', undefined, '文字转曲');
outlineCheck.value = true;

// 嵌入链接图选项
var embedCheck = dialog.add('checkbox', undefined, '嵌入链接图');
embedCheck.value = true;

// 按钮组
var buttonGroup = dialog.add('group');
buttonGroup.add('button', undefined, '确定', {name: 'ok'});
buttonGroup.add('button', undefined, '取消', {name: 'cancel'});

var copyright = dialog.add('statictext', undefined, 'By Hwangzhun');
embedCheck.value = true;

// 显示对话框
var result = dialog.show();

if (result == 1) {
    var doc = app.activeDocument;
    var doOutline = outlineCheck.value;
    var doEmbed = embedCheck.value;
    var suffix = suffixInput.text;
    
    // 1. 解锁所有图层和对象
    function unlockAllItems(container) {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            if (item.locked) item.locked = false;
        }
        if (container.layers) {
            for (var j = 0; j < container.layers.length; j++) {
                var layer = container.layers[j];
                if (layer.locked) layer.locked = false;
                unlockAllItems(layer);
            }
        }
    }
    unlockAllItems(doc);

    // 2. 嵌入所有链接图像（如果选项开启）
    if (doEmbed) {
        function embedAllLinkedImages(container) {
            for (var i = 0; i < container.pageItems.length; i++) {
                var item = container.pageItems[i];
                if (item.typename === "PlacedItem" && !item.embedded) {
                    item.embed();
                } else if (item.typename === "GroupItem") {
                    embedAllLinkedImages(item); // 递归嵌套
                }
            }
        }
        embedAllLinkedImages(doc);
    }

    // 3. 全选并创建轮廓（如果选项开启）
    if (doOutline) {
        app.executeMenuCommand("selectall");
        try {
            app.executeMenuCommand("outline"); // 创建轮廓
        } catch (e) {
            showAlert("创建轮廓时出错：" + e.message, "错误");
        }
    }

    // 4. 确定保存路径
    var originalName = doc.name.replace(/\.[^\.]+$/, '');
    var folder = doc.path;
    
    // 如果选择其他位置
    if (locationDropdown.selection.index == 1) {
        var newFolder = Folder.selectDialog("选择导出目录");
        if (newFolder === null) {
            showAlert("已取消选择导出位置，脚本终止执行。", "操作取消");
            exit(); // 终止脚本执行
        }
        folder = newFolder;
    }
    
    var pdfName = originalName + suffix + ".pdf";
    var pdfFile = new File(folder + "/" + pdfName);

    var pdfOptions = new PDFSaveOptions();
    pdfOptions.compatibility = PDFCompatibility.ACROBAT5;
    pdfOptions.preserveEditability = false;
    pdfOptions.artBoardClipping = false; // 关闭裁切，保存画板外的内容

    doc.saveAs(pdfFile, pdfOptions); // 保存为 PDF

    showAlert("脚本执行完成！文件已保存为: " + pdfFile.fsName, "完成");
} else {
    showAlert("已取消脚本执行。", "操作取消");
}

function exit() {
    // 空函数，用于终止脚本
    throw new Error("Script terminated by user");
}
