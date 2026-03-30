const { app, BrowserWindow, shell, dialog, Menu } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const http = require('http');

let mainWindow;
let serverProcess;
const PORT = 5000;

// 判断是否为开发环境
const isDev = !app.isPackaged;

// 检查端口是否可用
function checkServer(url) {
  return new Promise((resolve) => {
    http.get(url, (res) => {
      resolve(res.statusCode === 200);
    }).on('error', () => {
      resolve(false);
    });
  });
}

// 等待服务器启动
async function waitForServer(maxRetries = 30) {
  const url = `http://localhost:${PORT}`;
  for (let i = 0; i < maxRetries; i++) {
    if (await checkServer(url)) {
      return true;
    }
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  return false;
}

// 启动 Express 服务器
function startServer() {
  return new Promise((resolve, reject) => {
    const appPath = isDev 
      ? process.cwd()
      : path.dirname(app.getAppPath());
    
    console.log('Starting server from:', appPath);
    console.log('Is Dev:', isDev);

    serverProcess = spawn('node', ['server.js'], {
      cwd: appPath,
      shell: true,
      stdio: 'inherit',
      env: { ...process.env, PORT: String(PORT) }
    });

    serverProcess.on('error', (err) => {
      console.error('Failed to start server:', err);
      reject(err);
    });

    // 等待服务器启动
    waitForServer().then((started) => {
      if (started) {
        resolve();
      } else {
        reject(new Error('Server failed to start within timeout'));
      }
    });
  });
}

// 创建应用菜单
function createMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        { role: 'quit', label: '退出' }
      ]
    },
    {
      label: '视图',
      submenu: [
        { role: 'reload', label: '刷新' },
        { role: 'toggleDevTools', label: '开发者工具' },
        { type: 'separator' },
        { role: 'resetZoom', label: '重置缩放' },
        { role: 'zoomIn', label: '放大' },
        { role: 'zoomOut', label: '缩小' },
        { type: 'separator' },
        { role: 'togglefullscreen', label: '全屏' }
      ]
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '关于',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: '关于',
              message: '舱单文件处理系统',
              detail: '版本: 1.0.0\n自动生成提单确认件和装箱单发票'
            });
          }
        }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

// 创建主窗口
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    title: '舱单文件处理系统',
    show: false,
  });

  mainWindow.loadURL(`http://localhost:${PORT}`);

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('http://localhost') || url.startsWith('file://')) {
      return { action: 'allow' };
    }
    shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// 应用就绪
app.whenReady().then(async () => {
  try {
    createMenu();
    await startServer();
    createWindow();
  } catch (error) {
    console.error('Startup error:', error);
    dialog.showErrorBox('启动失败', `无法启动应用服务器: ${error.message}`);
    app.quit();
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// 关闭应用时停止服务器
app.on('window-all-closed', () => {
  if (serverProcess) {
    serverProcess.kill();
  }
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// 处理未捕获的异常
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  dialog.showErrorBox('错误', `应用发生错误: ${error.message}`);
});
