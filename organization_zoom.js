function(params) {
    if (!params.zoom || chart.__organizationFontUpdate) return;
    const series = chart.getOption().series[0];
    const zoom = series.zoom || 1;
    const sizes = {
        title: [18, 25], ceo: [22, 32], small: [10, 25],
        leader: [14, 23], member: [13, 22], rule: [11, 22]
    };
    const rich = {};
    for (const [name, dimensions] of Object.entries(sizes)) {
        rich[name] = {
            fontSize: dimensions[0] * zoom,
            lineHeight: dimensions[1] * zoom
        };
    }
    chart.__organizationFontUpdate = true;
    try {
        chart.setOption({series: [{id: series.id, label: {
            fontSize: 13 * zoom, lineHeight: 22 * zoom, rich: rich
        }}]}, {silent: true, lazyUpdate: false});
    } finally {
        chart.__organizationFontUpdate = false;
    }
    // undefined prevents a Streamlit rerun: zoom and pan stay in the browser.
}
