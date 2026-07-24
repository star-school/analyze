import os

def build_standalone():
    with open('analysis.html', 'r', encoding='utf-8') as f:
        html_content = f.read()

    with open('style.css', 'r', encoding='utf-8') as f:
        css_content = f.read()

    with open('analysis.js', 'r', encoding='utf-8') as f:
        js_content = f.read()

    # Inject CSS
    html_content = html_content.replace(
        '<link rel="stylesheet" href="style.css">',
        f'<style>\n{css_content}\n</style>'
    )

    # Replace JS Libraries with CDN links
    html_content = html_content.replace(
        '<script src="lib/xlsx.full.min.js"></script>',
        '<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>'
    )
    
    html_content = html_content.replace(
        '<script src="lib/chart.umd.js"></script>',
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js"></script>'
    )

    # Inject analysis.js
    html_content = html_content.replace(
        '<script src="analysis.js"></script>',
        f'<script>\n{js_content}\n</script>'
    )

    with open('入会対応分析ツール_軽量版.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

    print("Successfully built standalone HTML: 入会対応分析ツール_軽量版.html")

if __name__ == '__main__':
    build_standalone()
