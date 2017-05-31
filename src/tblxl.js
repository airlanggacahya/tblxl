
var tblxl = {}

tblxl.rgb2hex = function(colorval) {
    if (colorval === null || typeof(colorval) === "undefined") {
        return null
    }
    var parts = colorval.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/)
    if (parts === null) {
        return null;
    }
    parts.shift()
    for (var i = 0; i < 3; ++i) {
        parts[i] = parseInt(parts[i]).toString(16);
        if (parts[i].length == 1) parts[i] = '0' + parts[i];
    }
    return parts.join('');
}

tblxl.convertStyle = function(row, cell) {
    var style = {}

    // bg color
    v = tblxl.rgb2hex(cell.css("backgroundColor")) ||
        tblxl.rgb2hex(row.css("backgroundColor"))
    if (v)
        _.set(style, "fill.fgColor.rgb", "ff" + v)

    // font color
    v = tblxl.rgb2hex(cell.css("color"))
    if (v)
        _.set(style, "font.color.rgb", "ff" + v)

    // text align
    // start is left aligment
    v = cell.css("textAlign")
    if (v && v !== "start")
        _.set(style, "alignment.horizontal", v)

    // valign
    v = cell.css("verticalAlign")
    if (v) {
        switch (v) {
        case "middle": //convert middle to center
            _.set(style, "alignment.vertical", "center")
            break
        default:
            _.set(style, "alignment.vertical", v)
        }
    }

    var border = function(v){
        return {
            style: "thin",
            color: {
                rgb: "ff" + v
            }
        }
    }

    v = tblxl.rgb2hex(cell.css("borderTopColor"))
    if (v)
        _.set(style, "border.top", border(v))

    v = tblxl.rgb2hex(cell.css("borderRightColor"))
    if (v)
        _.set(style, "border.right", border(v))

    v = tblxl.rgb2hex(cell.css("borderBottomColor"))
    if (v)
        _.set(style, "border.bottom", border(v))

    v = tblxl.rgb2hex(cell.css("borderLeftColor"))
    if (v)
        _.set(style, "border.left", border(v))
    
    return style
}

tblxl.table2sheet = function(id) {
    var element = document.getElementById(id)
    var ws = XLSX.utils.table_to_sheet(element)
    ws["!rows"] = []
    ws["!cols"] = []

    $(element).find("tbody tr").each(function (rowidx, rowel) {
        var row = $(rowel)

        row.children("td").each(function (idx, cellel) {
            var cell = $(cellel)
            var top = cell.cellPos().top
            var left = cell.cellPos().left

            // overwrite width & height
            ws["!cols"][left] = {
                wpx: cellel.offsetWidth
            }
            // mark row contains cell
            ws["!rows"][top] = true

            var xlscell = ws[XLSX.utils.encode_cell({r:top, c:left})]
            xlscell.s = tblxl.convertStyle(row, cell)
        })
    })

    // set unmarked row as hidden row
    var ref = XLSX.utils.decode_range(ws["!ref"])
    for (var i = ref.s.r; i <= ref.e.r; i++) {
        if (ws["!rows"][i] == undefined) {
            ws["!rows"][i] = {
                hidden: true
            }
        } else {
            ws["!rows"][i] = undefined
        }
    }

    // handle merge cell border color
    _.each(ws["!merges"], function (range) {
        var r,c
        var src = XLSX.utils.encode_cell({r: range.s.r, c: range.s.c})
        for (var r = range.s.r; r <= range.e.r; r++) {
            for (var c = range.s.c; c <= range.e.c; c++) {
                if (r === range.s.r && c == range.s.c)
                    continue;

                var dst = XLSX.utils.encode_cell({r:r, c:c})
                _.set(ws, dst + ".s.border", _.get(ws, src + ".s.border"))
            }
        }
    })

    return ws
}

tblxl.save = function(ws, filename) {
    var wopts = {bookType:'xlsx', bookSST:false, type:'binary'};
    var wbout = XLSX.write({SheetNames: ["Sheet1"], Sheets: {Sheet1: ws}}, wopts);

    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    saveAs(new Blob([s2ab(wbout)], {type:"application/octet-stream"}), filename);
}

tblxl.setColumnWidth = function(ws, widths) {
    ws["!cols"] = _.map(widths, function (i) {
        return {
            wpx: i
        }
    })

    return ws
}