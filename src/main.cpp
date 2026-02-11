#include <QApplication>
#include <QMainWindow>
#include <QTranslator>
#include "cc/neolux/fem/xlsxeditor/XLSXEditor.hpp"

using namespace cc::neolux::fem::xlsxeditor;

int main(int argc, char *argv[]) {
    QApplication app(argc, argv);

    // Load translations
    QTranslator translator;
    if (translator.load("translations/XLSXEditor_zh_CN.qm")) {
        app.installTranslator(&translator);
    }

    // For testing, create a main window and embed XLSXEditor
    QMainWindow window;
    XLSXEditor *editor = new XLSXEditor(&window);
    editor->loadXLSX("../assets/test.xlsx", "SO13(DNo.3)MS", "B:K,7:34");
    window.setCentralWidget(editor);
    window.show();

    return app.exec();
}
