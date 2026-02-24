#include <QApplication>
#include <QMainWindow>
#include <QPushButton>
#include <QTimer>
#include <QTranslator>

#include "cc/neolux/fem/xlsxeditor/XLSXEditor.hpp"

using namespace cc::neolux::fem::xlsxeditor;

int main(int argc, char* argv[]) {
    // 启动参数：
    //   argv[1] = xlsx 文件路径
    //   --dry-run = 假删除模式（仅标记）
    //   --auto-test = 自动执行加载、标记、保存、退出流程
    if (argc < 2) {
        qWarning("Usage: %s <xlsx-file> [--dry-run]", argv[0]);
        qWarning("  --dry-run: Enable fake delete mode (default is real delete mode)");
        return -1;
    }
    std::string xlsxFile = argv[1];

    // 解析命令行参数
    bool dryRun = false;
    bool autoTest = false;
    for (int i = 2; i < argc; ++i) {
        if (std::string(argv[i]) == "--dry-run") {
            dryRun = true;
        } else if (std::string(argv[i]) == "--auto-test") {
            autoTest = true;
        }
    }

    QApplication app(argc, argv);

    // 加载翻译资源（供 FemApp 同步复用）
    QTranslator translator;
    if (translator.load("translations/XLSXEditor_zh_CN.qm")) {
        app.installTranslator(&translator);
    }

    // 测试入口：创建主窗口并嵌入 XLSXEditor 组件
    QMainWindow window;
    XLSXEditor* editor = new XLSXEditor(&window, dryRun);
    window.setCentralWidget(editor);
    window.resize(1024, 768);
    window.show();

    // 延迟加载，确保进度条可见
    QTimer::singleShot(0, editor, [editor, xlsxFile, autoTest]() {
        editor->loadXLSX(QString::fromStdString(xlsxFile), "SO13(DNo.3)MS", "B:K,7:34");

        if (autoTest) {
            QTimer::singleShot(500, editor, [editor]() {
                // 通过触发每个 DataItem 的按钮来批量切换删除状态
                QList<QWidget*> items = editor->findChildren<QWidget*>("dataItem");
                for (auto* item : items) {
                    // 模拟点击组件内按钮（用于自动化流程验证）
                    QList<QPushButton*> btns = item->findChildren<QPushButton*>();
                    if (!btns.isEmpty()) {
                        btns[0]->click();
                    }
                }

                // 标记完成后触发保存
                QTimer::singleShot(500, editor, [editor]() {
                    // 查找并点击保存按钮
                    QList<QPushButton*> saveBtns = editor->findChildren<QPushButton*>("btnSave");
                    if (!saveBtns.isEmpty()) {
                        saveBtns[0]->click();
                    }

                    // 保存后自动退出，便于脚本化验证
                    QTimer::singleShot(1000, []() { QApplication::quit(); });
                });
            });
        }
    });

    return app.exec();
}
