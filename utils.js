// utils.js

function setCellStyle(ri, ci, property, value, sheetIndex = 0) {
    const dataProxy = s.datas[sheetIndex];
    const { styles, rows } = dataProxy;
    const cell = rows.getCellOrNew(ri, ci);
    let cstyle = {};

    if (cell.style !== undefined) {
        cstyle = { ...styles[cell.style] };
    }
    if (property.startsWith('font')) {
        const nfont = {};
        nfont[property.split('-')[1]] = value;
        cstyle.font = { ...(cstyle.font || {}), ...nfont };
    } else {
        cstyle[property] = value;
    }
    cell.style = dataProxy.addStyle(cstyle);
    s.reRender();
}


