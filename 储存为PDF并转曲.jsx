if (confirm("是否执行转曲脚本？")) {
    var doc = app.activeDocument;

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

    // 2. 嵌入所有链接图像
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

    // 3. 全选并创建轮廓
    app.executeMenuCommand("selectall");
    try {
        app.executeMenuCommand("outline"); // 创建轮廓
    } catch (e) {
        alert("创建轮廓时出错：" + e.message);
    }

    // 4. 保存修改
    doc.save();

    // 5. 储存为 PDF（加 -转曲 后缀，不弹窗）
    var originalName = doc.name.replace(/\.[^\.]+$/, '');
    var folder = doc.path;
    var pdfName = originalName + "-转曲.pdf";
    var pdfFile = new File(folder + "/" + pdfName);

    var pdfOptions = new PDFSaveOptions();
    pdfOptions.compatibility = PDFCompatibility.ACROBAT5;
    pdfOptions.preserveEditability = false;
    pdfOptions.artBoardClipping = false; // 关闭裁切，保存画板外的内容

    doc.saveAs(pdfFile, pdfOptions); // 保存为 PDF

    // 6. 可选：关闭 PDF 文档（如果仍然开启）
    //doc.close(SaveOptions.DONOTSAVECHANGES); // 关闭 .ai 文件，无需保存

    alert("脚本执行完成！");
} else {
    alert("已取消脚本执行。");
}
