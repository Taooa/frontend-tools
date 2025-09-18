// 全局数据存储
window.dataSources = {};
let currentPreviewKey = null;
let currentFile = null;

// 在文件开头的DOM元素定义之后，添加init函数
// 更新DOM元素获取部分
// DOM 元素
// 删除重复定义的fileInput变量
const fileList = document.getElementById('fileList');
const dataPreview = document.getElementById('dataPreview');
const codeInput = document.getElementById('codeInput');
const runCodeBtn = document.getElementById('runCodeBtn');
const codeResult = document.getElementById('codeResult');
const exportDataSource = document.getElementById('exportDataSource');
const exportBtn = document.getElementById('exportBtn');
const uploadBtn = document.getElementById('uploadBtn');
const uploadModal = document.getElementById('uploadModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const cancelUploadBtn = document.getElementById('cancelUploadBtn');
const confirmUploadBtn = document.getElementById('confirmUploadBtn');
const dataSourceName = document.getElementById('dataSourceName');
const hasHeader = document.getElementById('hasHeader');
const currentFileName = document.getElementById('currentFileName');
const nameError = document.getElementById('nameError');
// 添加新的DOM元素
const fileUploadInModal = document.getElementById('fileUploadInModal');
const selectFileBtn = document.getElementById('selectFileBtn');

// 添加缺失的init函数定义
function init() {
  setupEventListeners();
  updateFileList();
  updateExportOptions();
}

// 修改setupEventListeners函数
function setupEventListeners() {
  // 文件上传按钮点击 - 直接显示弹窗
  uploadBtn.addEventListener('click', () => {
    // 生成默认的数据源名称
    const defaultName = `data${Object.keys(window.dataSources).length + 1}`;
    dataSourceName.value = defaultName;
    currentFileName.textContent = '未选择文件';
    nameError.textContent = '';
    currentFile = null;
    // 移除对fileInput的引用，因为我们现在使用fileUploadInModal
    fileUploadInModal.value = '';

    // 显示弹窗
    uploadModal.style.display = 'block';
  });

  // 弹窗中的文件选择按钮点击
  selectFileBtn.addEventListener('click', () => {
    fileUploadInModal.click();
  });

  // 文件选择变化 - 使用新的file input元素
  fileUploadInModal.addEventListener('change', (event) => {
    if (event.target.files.length > 0) {
      currentFile = event.target.files[0];
      if (!currentFile.name.endsWith('.xlsx') && !currentFile.name.endsWith('.xls')) {
        alert('请上传Excel文件(.xlsx或.xls)');
        currentFile = null;
        currentFileName.textContent = '未选择文件';
        return;
      }

      currentFileName.textContent = currentFile.name;
    }
  });

  // 关闭弹窗
  closeModalBtn.addEventListener('click', closeModal);
  cancelUploadBtn.addEventListener('click', closeModal);

  // 点击弹窗外部关闭
  window.addEventListener('click', (event) => {
    if (event.target === uploadModal) {
      closeModal();
    }
  });

  // 确认上传
  confirmUploadBtn.addEventListener('click', processFileWithOptions);

  // 数据源名称输入验证
  dataSourceName.addEventListener('input', () => {
    validateDataSourceName();
  });

  // 运行代码
  runCodeBtn.addEventListener('click', runCode);

  // 导出数据
  exportBtn.addEventListener('click', exportData);
}

// 更新processFileWithOptions函数，添加文件验证
function processFileWithOptions() {
  if (!validateDataSourceName() || !currentFile) {
    if (!currentFile) {
      alert('请选择要上传的Excel文件');
    }
    return;
  }

  const key = dataSourceName.value.trim();
  const includeHeader = hasHeader.checked;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // 获取工作表的范围
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      const jsonData = [];

      // 根据是否包含表头确定起始行
      const startRow = includeHeader ? range.s.r + 1 : range.s.r;

      // 处理每一行数据
      for (let R = startRow; R <= range.e.r; R++) {
        const row = {};
        let hasData = false;

        // 为每一列创建field1, field2...的属性名
        for (let C = range.s.c; C <= range.e.c; C++) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = worksheet[cellAddress];
          const cellValue = cell ? cell.v : '';

          if (cellValue !== '') {
            hasData = true;
          }

          // 始终使用field1, field2...格式的属性名
          const fieldName = `field${C + 1}`;
          row[fieldName] = cellValue;
        }

        // 只添加有数据的行
        if (hasData) {
          jsonData.push(row);
        }
      }

      // 保存数据
      window.dataSources[key] = jsonData;

      // 更新UI
      updateFileList();
      updateExportOptions();

      // 如果是第一个文件，自动预览
      if (Object.keys(window.dataSources).length === 1) {
        renderDataPreview(key);
        currentPreviewKey = key;
      }

      // 关闭弹窗
      closeModal();

    } catch (error) {
      alert(`处理文件 ${currentFile.name} 时出错: ${error.message}`);
    }
  };

  reader.readAsArrayBuffer(currentFile);
}

// 更新closeModal函数，清空新的file input
function closeModal() {
  uploadModal.style.display = 'none';
  currentFile = null;
  // 注意：这里还引用了fileInput，但我们已经删除了它，需要修复
  // fileInput.value = '';
  fileUploadInModal.value = '';
}

// 验证数据源名称
function validateDataSourceName() {
  const name = dataSourceName.value.trim();
  const regex = /^[a-zA-Z0-9]+$/;

  if (!name) {
    nameError.textContent = '请输入数据源名称';
    return false;
  }

  if (!regex.test(name)) {
    nameError.textContent = '名称只能包含英文和数字';
    return false;
  }

  if (window.dataSources.hasOwnProperty(name)) {
    nameError.textContent = '该名称已存在，请使用其他名称';
    return false;
  }

  nameError.textContent = '';
  return true;
}

// 更新文件列表
function updateFileList() {
  fileList.innerHTML = '';

  for (const key in window.dataSources) {
    const fileItem = document.createElement('div');
    fileItem.className = 'file-item';
    fileItem.innerHTML = `
            <div>
                <span>${key}</span> - ${window.dataSources[key].length} 条记录
                <button class="preview-btn" data-key="${key}">预览</button>
            </div>`;

    // 添加预览按钮点击事件
    const previewBtn = fileItem.querySelector('.preview-btn');
    previewBtn.addEventListener('click', () => {
      renderDataPreview(key);
      currentPreviewKey = key;
    });

    fileList.appendChild(fileItem);
  }
}

// 更新导出选项
function updateExportOptions() {
  exportDataSource.innerHTML = '<option value="">选择要导出的数据源</option>';

  for (const key in window.dataSources) {
    const option = document.createElement('option');
    option.value = key;
    option.textContent = `${key} (${window.dataSources[key].length} 条记录)`;
    exportDataSource.appendChild(option);
  }
}

// 渲染数据预览
function renderDataPreview(key) {
  const data = window.dataSources[key];
  if (!data || data.length === 0) {
    dataPreview.innerHTML = '<p>没有数据可预览</p>';
    return;
  }

  const headers = Object.keys(data[0]);

  let html = `
        <h3>${key} (${data.length} 条记录)</h3>
        <table class="data-table">
            <thead>
                <tr>`;

  // 添加表头
  headers.forEach(header => {
    html += `<th>${header}</th>`;
  });

  html += `
                </tr>
            </thead>
            <tbody>`;

  // 添加数据行（最多显示100行）
  const displayData = data.slice(0, 100);
  displayData.forEach(row => {
    html += '<tr>';
    headers.forEach(header => {
      html += `<td>${row[header] !== undefined ? row[header] : ''}</td>`;
    });
    html += '</tr>';
  });

  if (data.length > 100) {
    html += `<tr><td colspan="${headers.length}" style="text-align: center; color: #666;">仅显示前100条记录，共 ${data.length} 条</td></tr>`;
  }

  html += `
            </tbody>
        </table>`;

  dataPreview.innerHTML = html;
}

// 运行代码
function runCode() {
  codeResult.innerHTML = '';

  try {
    // 捕获控制台输出
    const originalConsoleLog = console.log;
    const logs = [];
    console.log = (...args) => {
      logs.push(args.map(arg => {
        if (typeof arg === 'object') {
          return JSON.stringify(arg, null, 2);
        }
        return arg;
      }).join(' '));
      originalConsoleLog.apply(console, args);
    };

    // 执行用户代码
    new Function(codeInput.value)();

    // 恢复控制台
    console.log = originalConsoleLog;

    // 显示执行结果
    if (logs.length > 0) {
      codeResult.innerHTML = logs.join('<br>');
    } else {
      codeResult.innerHTML = '代码执行成功！';
    }

    // 更新文件列表和导出选项（以防用户添加了新的数据）
    updateFileList();
    updateExportOptions();

  } catch (error) {
    codeResult.innerHTML = `执行错误: ${error.message}`;
    codeResult.style.color = '#e74c3c';
  }
}

// 导出数据
function exportData() {
  const key = exportDataSource.value;
  if (!key || !window.dataSources[key]) {
    alert('请选择要导出的数据源');
    return;
  }

  try {
    const data = window.dataSources[key];
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, key);

    // 生成文件名
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const fileName = `${key}_${timestamp}.xlsx`;

    // 导出文件
    XLSX.writeFile(workbook, fileName);

    codeResult.innerHTML = `数据已成功导出为 ${fileName}`;
  } catch (error) {
    alert(`导出数据时出错: ${error.message}`);
  }
}

// 页面加载完成后初始化
window.addEventListener('load', init);