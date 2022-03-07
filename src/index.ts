import {
  QMainWindow,
  QWidget,
  FlexLayout,
  QPushButton,
  FileMode,
  QFileDialog,
  QApplication,
  NodeWidget,
} from '@nodegui/nodegui';
import xlsx from 'node-xlsx';
import storage from 'node-persist';

const win = new QMainWindow();
win.setWindowTitle('XLSX Converter');

let startupFolder;
let fileDialog: QFileDialog;

const boot = async () => {
  await storage.init();
  startupFolder = await storage.getItem('folder');
  fileDialog = new QFileDialog(win, 'Select file', startupFolder);
  fileDialog.setFileMode(FileMode.AnyFile);
  fileDialog.setNameFilter('Excel (*.xlsx)');
  win.show();
}

const centralWidget = new QWidget();
centralWidget.setObjectName('myroot');
const rootLayout = new FlexLayout();
centralWidget.setLayout(rootLayout);

const button = new QPushButton();
button.setText('Click me');

const clipboard = QApplication.clipboard();

button.addEventListener('clicked', async() => {
  const lastFolder = await storage.getItem('folder');
  fileDialog = new QFileDialog(win, 'Select file', lastFolder);
  fileDialog.exec();
  const selectedFiles = fileDialog.selectedFiles();
  selectedFiles.map(async (file) => {
    const splittedFolder = file.split('/');
    splittedFolder.pop();
    const folder = splittedFolder.join('/');
    await storage.setItem('folder',folder)
    try {
      const resultArr = xlsx.parse(file);
      resultArr.map(({ data }) => {
        const reducedText: any = data.reduce((acc: any, item: any) => {
          if (!item.length || !parseInt(item[1], 10)) return acc;
          acc.push(`${item[1]}:true`);
          return acc;
        }, []);
        clipboard.setText(`${reducedText.join(';')};`);
      });
    } catch (err) {
      console.error(err);
    }
  });
});

rootLayout.addWidget(button);
win.setCentralWidget(centralWidget);
win.setStyleSheet(
  `
    #myroot {
      background-color: #009688;
      height: '100%';
      align-items: 'center';
      justify-content: 'center';
    }
    #mylabel {
      font-size: 16px;
      font-weight: bold;
      padding: 1;
    }
  `
);
boot();

(global as any).win = win;
