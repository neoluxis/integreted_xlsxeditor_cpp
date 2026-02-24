#include <QApplication>
#include <QMainWindow>
#include <QTimer>
#include <QTranslator>

#include "cc/neolux/fem/xlsxeditor/XLSXEditor.hpp"

using namespace cc::neolux::fem::xlsxeditor;

int main(int argc, char* argv[]) {
    if (argc < 2) {
        qWarning("Usage: %s <xlsx-file>", argv[0]);
        return -1;
    }
    std::string xlsxFile = argv[1];

    QApplication app(argc, argv);

    // Load translations
    QTranslator translator;
    if (translator.load("translations/XLSXEditor_zh_CN.qm")) {
        app.installTranslator(&translator);
    }

    // For testing, create a main window and embed XLSXEditor
    QMainWindow window;
    XLSXEditor* editor = new XLSXEditor(&window);
    window.setCentralWidget(editor);
    window.resize(1024, 768);
    window.show();

    // 延迟加载，确保进度条可见
    QTimer::singleShot(0, editor, [editor, xlsxFile]() {
        editor->loadXLSX(QString::fromStdString(xlsxFile), "SO13(DNo.3)MS", "B:K,7:34");
    });

    return app.exec();
}
