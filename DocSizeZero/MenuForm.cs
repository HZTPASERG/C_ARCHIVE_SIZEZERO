using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Windows.Media.Imaging;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace DocSizeZero
{
    public partial class MenuForm : Form
    {
        private MainApp _mainApp;
        private DatabaseService _databaseService;
        private ImageList _imageList;
        private Panel previewPanel;
        public DataTable documentTable;
        private bool _isClosing = false;
        private List<string> tempFiles = new List<string>();    // Lista de ficheros temporales que muestran durante la sesion

        public MenuForm(MainApp mainApp)
        {
            InitializeComponent();
            _mainApp = mainApp ?? throw new ArgumentNullException(nameof(mainApp));

            // Inicializar el servicio de base de datos con la cadena de conexión
            _databaseService = new DatabaseService(_mainApp.DatabaseManagerPersistent);

            // Configurar TreeView para un estilo más moderno
            ConfigureTreeView();

            // Inicializa el panel de vista previa
            InitializePreviewPanel();

            // Asocia el evento AfterSelect
            treeView.AfterSelect += treeView_AfterSelect;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Cargar los documentos en la tabla "documentTable"
            documentTable = _databaseService.LoadDocumentList(); // Carga todos los documentos

            // Llama al método para cargar los datos
            LoadTreeView();
        }

        private void LoadTreeView()
        {
            try
            {
                // Cargar los datos desde la base de datos
                List<TreeNodeModel> nodes = _databaseService.LoadTreeData(documentTable);

                if (nodes == null || nodes.Count == 0)
                {
                    MessageBox.Show("No se encontraron datos para el TreeView.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Obtener las claves únicas de las imágenes desde los nodos cargados
                var imageKeys = nodes
                    .Select(n => n.ImageId) // Asumiendo que `ImageId` es el identificador de la imagen en la BD
                    .Distinct()
                    .ToList();

                if (imageKeys.Count == 0)
                {
                    MessageBox.Show("No se encontraron claves de imágenes asociadas con los nodos del TreeView.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Cargar imágenes desde la base de datos usando las claves únicas
                // Después de cargar los datos del árbol, cargamos las imágenes y configúramoslas en el TreeView
                var images = _databaseService.LoadImageList(imageKeys);
                InitializeImageList(images);

                // Rellenar el TreeView con los datos
                PopulateTreeView(nodes);

                // Manejar el evento BeforeExpand para cargar subnodos y documentos dinámicamente
                treeView.BeforeExpand += TreeView_BeforeExpand;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar el TreeView: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PopulateTreeView(List<TreeNodeModel> nodes)
        {
            treeView.Nodes.Clear(); // Limpiar nodos existentes

            // Encontrar los nodos raíz (ParentId == "ROOT_0")
            var rootNodes = nodes.Where(n => n.ParentId == "ROOT_0").ToList();

            foreach (var rootNode in rootNodes)
            {
                TreeNode treeNode = new TreeNode
                {
                    Text = rootNode.Name,
                    Tag = rootNode, // Aquí se debe asignar el objeto TreeNodeModel
                    ImageKey = rootNode.ImageId.ToString(), // Asignar imagen por clave
                    SelectedImageKey = rootNode.ImageId.ToString()
                };

                // AddChildNodes(treeNode, rootNode.Id, nodes);

                // Si el nodo tiene hijos, agregar un placeholder
                if (rootNode.HasChildren)
                {
                    treeNode.Nodes.Add(new TreeNode("Cargando..."));
                }

                treeView.Nodes.Add(treeNode);
            }
        }

        private void AddChildNodes(TreeNode parentNode, string parentId, List<TreeNodeModel> nodes)
        {
            var childNodes = nodes.Where(n => n.ParentId == parentId).ToList();

            foreach (var childNode in childNodes)
            {
                TreeNode treeNode = new TreeNode
                {
                    Text = childNode.Name,
                    Tag = childNode.Id,
                    ImageKey = childNode.ImageId.ToString(), // Asignar imagen por clave
                    SelectedImageKey = childNode.ImageId.ToString()
                };

                AddChildNodes(treeNode, childNode.Id, nodes);
                parentNode.Nodes.Add(treeNode);
            }
        }

        private void TreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode parentNode = e.Node;
            //string parentId = parentNode.Tag.ToString();
            string parentId = parentNode.Tag is TreeNodeModel model ? model.Id : null;

            // Limpiar subnodos si ya se cargaron previamente
            if (!string.IsNullOrEmpty(parentId) && parentNode.Nodes.Count == 1 && parentNode.Nodes[0].Text == "Cargando...")
            {
                parentNode.Nodes.Clear();
    Console.WriteLine($"parentNode.Tag: {parentId}");
                // Cargar los nodos hijos y documentos relacionados
                var childNodes = _databaseService.LoadChildNodes(parentId); // Nodos hijos
                var documents = _databaseService.GetDocumentsForNode(parentId, documentTable); // Documentos

                // Obtener claves de imágenes de los documentos y cargar dinámicamente
                var imageKeys = documents.Select(d => d.ImageId).Distinct().ToList();
                if (imageKeys.Count > 0)
                {
                    var newImages = _databaseService.LoadImageList(imageKeys);
                    foreach (var kvp in newImages)
                    {
                        if (!_imageList.Images.ContainsKey(kvp.Key.ToString()))
                        {
                            try
                            {
                                _imageList.Images.Add(kvp.Key.ToString(), kvp.Value);
                                Console.WriteLine($"Imagen añadida dinámicamente para la clave: {kvp.Key}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error al agregar la imagen para la clave {kvp.Key}: {ex.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Imagen con clave {kvp.Key} ya existe en el ImageList.");
                        }
                    }
                }

                // Agregar nodos hijos al nodo expandido
                foreach (var childNode in childNodes)
                {
                    TreeNode childTreeNode = new TreeNode
                    {
                        Text = childNode.Name,
                        Tag = childNode,    // Guardar el TreeNodeModel completo en el Tag
                        ImageKey = childNode.ImageId.ToString(),
                        SelectedImageKey = childNode.ImageId.ToString()
                    };

                    // Marcar visualmente si es día u hora con diferencia
                    if ((childNode.Table == "DAY" && childNode.DiaRes == 1) ||
                        (childNode.Table == "HOUR" && childNode.HoraRes == 1))
                    {
                        childTreeNode.ForeColor = Color.Red;
                        childTreeNode.NodeFont = new Font(treeView.Font, FontStyle.Bold);
                    }

                    // Si el nodo puede tener subnodos, añadir un placeholder "Cargando..."
                    if (childNode.HasChildren)
                    {
                        childTreeNode.Nodes.Add(new TreeNode("Cargando..."));
                    }

                    parentNode.Nodes.Add(childTreeNode);
                }

                // Agregar documentos como subnodos
                foreach (var doc in documents)
                {
                    TreeNode documentNode = new TreeNode
                    {
                        Text = doc.Name,
                        Tag = doc,  // Guardar el TreeNodeModel completo en el Tag
                        ImageKey = doc.ImageId.ToString(),
                        SelectedImageKey = doc.ImageId.ToString()
                    };

                    // Marcar documentos que no existwn en la fecha anterior
                    if (doc.IsNew == 1)
                    {
                        documentNode.ForeColor = Color.Red;
                        documentNode.NodeFont = new Font(treeView.Font, FontStyle.Bold);
                    }

                    parentNode.Nodes.Add(documentNode);
                }
            }
        }

        // Añade una ImageList para asociar las imágenes con los nodos del TreeView
        private void InitializeImageList(Dictionary<int, Image> images)
        {
            try
            {
                //  Limpiar cualquier instancia previa
                if (_imageList != null)
                {
                    _imageList.Dispose();
                    _imageList = null;
                }

                _imageList = new ImageList();
                _imageList.ImageSize = new Size(32, 32);

                Console.WriteLine($"Total de imágenes a procesar: {images.Count}");

                foreach (var kvp in images)
                {
                    if (kvp.Value == null)
                    {
                        Console.WriteLine($"Imagen nula para la clave: {kvp.Key}");
                        continue;
                    }

                    try
                    {
                        /*
                        if (!IsImageValid(kvp.Value))
                        {
                            Console.WriteLine($"Imagen inválida para la clave: {kvp.Key}");
                            continue;
                        }
                        */

                        _imageList.Images.Add(kvp.Key.ToString(), ResizeImage(kvp.Value, _imageList.ImageSize));
                        // _imageList.Images.Add(kvp.Key.ToString(), kvp.Value);
                        Console.WriteLine($"Imagen añadida para la clave: {kvp.Key}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al añadir la imagen para la clave {kvp.Key}: {ex.Message}");
                    }
                }

                treeView.ImageList = _imageList;
                Console.WriteLine("ImageList inicializado correctamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al inicializar el ImageList: {ex.Message}");
                MessageBox.Show($"Error al inicializar el ImageList: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Establece un tamaño predeterminado para la imagen
        private Image ResizeImage(Image img, Size size)
        {
            Bitmap bmp = new Bitmap(size.Width, size.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.DrawImage(img, new Rectangle(0, 0, size.Width, size.Height));
            }
            return bmp;
        }

        // Validar la imagen
        private bool IsImageValid(Image img)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    img.Save(ms, img.RawFormat);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            Console.WriteLine($"Nodo seleccionado: {e.Node.Text}");

            TreeNode selectedNode = e.Node;

            // Verificar si el Tag del nodo es un TreeNodeModel
            if (selectedNode.Tag is TreeNodeModel nodeModel)
            {
                // Verificar si el nodo pertenece a la tabla "DOCUMENT"
                if (nodeModel.Table == "DOCUMENT")
                {
                    int docId = nodeModel.Key;

                    // Obtener detalles del documento usando el Key y Table
                    DocumentModel document = _databaseService.GetDocumentDetails(docId);

                    if (document == null || document.FileBody.Length == 0)
                    {
                        MessageBox.Show($"No se pudo recuperar el archivo con doc_id [{docId}] desde la base de datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        // Mostrar vista previa
                        ShowPreview(document, selectedNode);
                        treeView.SelectedNode = e.Node;
                    }
                }
            }
        }

        // Crear el panel en tiempo de ejecución
        private void InitializePreviewPanel()
        {
            previewPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.WhiteSmoke
            };

            // Agregar el panel al formulario
            splitContainer1.Panel2.Controls.Add(previewPanel);  // Asegura que el panel esté en el lugar correcto
        }

        // Mostrar el contenido del documento indicado en el controlador de la vista previa
        private void ShowPreview(DocumentModel document, TreeNode selectedNode)
        {
            // Obtenemos una ruta absoluta y verificamos si es válida utilizando Path.GetFullPath
            string fullPath;
            try
            {
                string filePath = Path.Combine(document.DirectoryPath, document.FileName);
                fullPath = Path.GetFullPath(filePath); // Verifica que la ruta sea válida
                Console.WriteLine($"Ruta completa: {fullPath}");

                if (!File.Exists(fullPath))
                {
                    // Si el archivo no existe, intenta recuperarlo
                    if (!RetrieveFileFromDatabase(document, fullPath))
                    {
                        ShowUnsupportedMessage($"El archivo [{fullPath}] no se encuentra disponible o no pudo ser descargado.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en la ruta: {ex.Message}");
                return;
            }

            // Mostrar vista previa
            string extension = Path.GetExtension(document.FileName).ToUpper();

            // Limpiar el panel de vista previa
            previewPanel.Controls.Clear();

            // Crear un contenedor principal (TableLayoutPanel)
            TableLayoutPanel layoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2, // Una columna para el toolbar y otra para el mainPanel
                RowCount = 1,
            };

            // Establecer tamaños proporcionales
            layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Auto para el toolbar
            layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // Resto para el mainPanel

            // Crear el panel principal para contener la vista previa y la barra de herramientas
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true, // Permite el desplazamiento
            };

            // Generar el ToolBar de la vista previa
            ToolStrip toolbar = ShowToolBarPreview(extension, fullPath, selectedNode, mainPanel);

            layoutPanel.Controls.Add(toolbar, 0, 0); // Agregar toolbar a la primera columna
            layoutPanel.Controls.Add(mainPanel, 1, 0); // Agregar mainPanel a la segunda columna

            previewPanel.Controls.Add(layoutPanel); // Agregar el layoutPanel al previewPanel

            switch (extension)
            {
                case ".PDF":
                    ShowPdfPreview(fullPath, mainPanel);
                    break;

                case ".JPG":
                    ShowImagePreview(fullPath, mainPanel);
                    break;
                case ".TIF":
                case ".TIFF":
                    LoadTiffWithWIC(fullPath, toolbar, mainPanel);
                    break;

                case "TXT":
                    ShowTextPreview(fullPath, mainPanel);
                    break;
                case ".DOC":
                case ".DOCX":
                    ShowWordPreview(fullPath, mainPanel);
                    break;

                default:
                    ShowUnsupportedMessage($"Tipo de archivo no soportado: {extension}");
                    break;
            }

            Console.WriteLine($"Control agregado: {previewPanel.Controls[0].GetType().Name}");
            Console.WriteLine($"Control visible: {previewPanel.Controls[0].Visible}");
            Console.WriteLine($"Panel dimensiones: {previewPanel.Width}x{previewPanel.Height}");
        }

        // Generar ToolBar de la vista previa
        private ToolStrip ShowToolBarPreview(string extension, string fullPath, TreeNode selectedNode, Panel mainPanel)
        {
            // Ruta a la carpeta de imágenes
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images");

            // Crear barra de herramientas
            ToolStrip toolbar = new ToolStrip
            {
                Dock = DockStyle.Left,
                GripStyle = ToolStripGripStyle.Hidden,
                BackColor = Color.Transparent,
                LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow,   // Estilo vertical
                AutoSize = true,
                Margin = new Padding(10, 10, 10, 10) // Margen externo: 10px alrededor del ToolStrip
            };

            // Crear botón "Reload"
            ToolStripButton btnReload = new ToolStripButton
            {
                Text = "",
                DisplayStyle = ToolStripItemDisplayStyle.Image,
                ToolTipText = "Завантажити документ повторно",
                Image = Image.FromFile(Path.Combine(imagePath, "reload.png")),
                Margin = new Padding(5, 5, 5, 5),   // Espacio libre alrededor del botón
                ImageScaling = ToolStripItemImageScaling.None,
                AutoSize = false,
                Width = 36,
                Height = 36
            };

            // Botón de abrir el documento en la aplicación predeterminada
            ToolStripButton btnRun = new ToolStripButton
            {
                Text = "Run",   // Opcional, por accesibilidad
                DisplayStyle = ToolStripItemDisplayStyle.Image,   // Solo mostrar la imagen
                ToolTipText = "Відукрити документ у власному додатку",
                Margin = new Padding(5, 5, 5, 5),   // Espacio libre alrededor del botón
            };
            // Asignar ícono basado en el tipo de documento
            if (selectedNode != null && !string.IsNullOrEmpty(selectedNode.ImageKey) && _imageList.Images.ContainsKey(selectedNode.ImageKey))
            {
                // Redimensionar la imagen a 32x32
                Image originalImage = _imageList.Images[selectedNode.ImageKey];
                Bitmap resizedImage = new Bitmap(originalImage, new Size(32, 32));
                btnRun.Image = resizedImage;    // Asignar la imagen redimensionada
            }
            // Ajustar el tamaño del botón al tamaño de la imagen
            btnRun.ImageScaling = ToolStripItemImageScaling.None;   // Evita que el botón escale la imagen
            btnRun.AutoSize = false; // Permite ajustar el tamaño del botón manualmente
            btnRun.Width = 36;
            btnRun.Height = 36;

            // Asignar eventos para cambiar el cursor
            // btnRun
            btnRun.MouseEnter += (s, e) => { Cursor.Current = Cursors.Hand; };  // Cambia a manoPrevious
            btnRun.MouseLeave += (s, e) => { Cursor.Current = Cursors.Default; };  // Restaura el cursor

            // btnReload
            btnReload.MouseEnter += (s, e) => { Cursor.Current = Cursors.Hand; };  // Cambia a manoPrevious
            btnReload.MouseLeave += (s, e) => { Cursor.Current = Cursors.Default; };  // Restaura el cursor

            // Agregar botones al panel
            toolbar.Items.Add(btnRun);
            toolbar.Items.Add(btnReload);

            // RUN
            btnRun.Click += (s, e) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(fullPath); // Abre el archivo en la app predeterminada
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Помилка вдкриття документа: {ex.Message}", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            
            // RELOAD
            btnReload.Click += (s, e) =>
            {
                try
                {
                    if (selectedNode?.Tag is TreeNodeModel nodeModel && nodeModel.Table == "DOCUMENT")
                    {
                        // Forzar la recarga del documento desde la base de datos
                        var document = _databaseService.ForceReloadDocument(nodeModel.Key);

                        if (document != null && document.FileBody.Length > 0)
                        {
                            string filePath = Path.Combine(document.DirectoryPath, document.FileName);

                            // Asegurarse de que el archivo no está en uso
                            if (IsFileLocked(filePath))
                            {
                                MessageBox.Show("Архів продовжує використовуватися іншим процесом.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            // Sobrescribir el archivo en el disco
                            File.WriteAllBytes(filePath, document.FileBody);

                            // Eliminar el contenido del mainPanel
                            foreach (Control control in mainPanel.Controls)
                            {
                                if (control is WebBrowser webBrowser)
                                {
                                    webBrowser.Navigate("about:blank"); // Navegar a una página en blanco
                                    webBrowser.Dispose(); // Liberar recursos del control
                                }
                            }
                            mainPanel.Controls.Clear();
                            
                            // Actualizar la vista previa
                            ShowPreview(document, selectedNode);
                        }
                        else
                        {
                            MessageBox.Show($"Помилка завантаження документа з бази данних.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Помилка завантаження документа: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            

            /*
            btnReload.Click += (s, e) =>
            {
                try
                {
                    if (File.Exists(fullPath))
                    {
                        // Eliminar el archivo actual del disco
                        File.Delete(fullPath);
                    }

                    // Forzar el evento treeView_AfterSelect para el nodo actual
                    treeView_AfterSelect(this, new TreeViewEventArgs(selectedNode));
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al recargar el documento: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            */

            // Agregar la barra de herramientas al panel de vista previa
            previewPanel.Controls.Add(toolbar);

            return toolbar;
        }

        // Recuperar el fichero del documento indicado
        private bool RetrieveFileFromDatabase(DocumentModel document, string fullPath)
        {
            try
            {
                // Crea el directorio si no existe
                string directoryPath = document.DirectoryPath;
                if (!Directory.Exists(document.DirectoryPath))
                {
                    Directory.CreateDirectory(document.DirectoryPath);
                    Console.WriteLine($"Directorio creado: {directoryPath}");
                }

                if (document.FileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    throw new ArgumentException("El nombre del archivo contiene caracteres inválidos.");                   }

                // Guarda el archivo en la ruta indicada
                File.WriteAllBytes(fullPath, document.FileBody);

                Console.WriteLine($"Archivo guardado en: {document.DirectoryPath}");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al recuperar el archivo desde la base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // PDF
        private void ShowPdfPreview(string fullPath, Panel mainPanel)
        {
            try
            {
                WebBrowser pdfPreview = new WebBrowser
                {
                    Dock = DockStyle.Fill
                };

                if (File.Exists(fullPath))
                {
                    pdfPreview.Navigate(fullPath);
                    mainPanel.Controls.Add(pdfPreview);
                    previewPanel.Refresh();
                }
                else
                {
                    pdfPreview.Navigate(fullPath);
                    mainPanel.Controls.Add(pdfPreview);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al mostrar PDF: {ex.Message}");
            }
        }

        // IMG
        private void ShowImagePreview(string fullPath, Panel mainPanel)
        {
            PictureBox imagePreview = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = Image.FromFile(fullPath)
            };
            mainPanel.Controls.Add(imagePreview);
        }

        //TXT
        private void ShowTextPreview(string fullPath, Panel mainPanel)
        {
            TextBox txtPreview = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Text = File.ReadAllText(fullPath)
            };
            mainPanel.Controls.Add(txtPreview);
        }

        // TIFF
        private void ShowTiffPreview(string fullPath)
        {
            try
            {
                Image tiffImage = Image.FromFile(fullPath);

                if (tiffImage.FrameDimensionsList.Length > 0)
                {
                    var dimension = new System.Drawing.Imaging.FrameDimension(tiffImage.FrameDimensionsList[0]);
                    int frameCount = tiffImage.GetFrameCount(dimension);

                    if (frameCount > 1)
                    {
                        // Mostrar un selector de página si el TIFF tiene múltiples páginas
                        Panel multiPagePanel = new Panel
                        {
                            Dock = DockStyle.Fill
                        };

                        PictureBox imagePreview = new PictureBox
                        {
                            Dock = DockStyle.Fill,
                            SizeMode = PictureBoxSizeMode.Zoom,
                            Image = GetTiffFrame(tiffImage, 0) // Mostrar la primera página
                        };

                        Button nextButton = new Button
                        {
                            Text = "Siguiente",
                            Dock = DockStyle.Bottom
                        };
                        nextButton.Click += (s, e) =>
                        {
                            int currentFrame = int.Parse(nextButton.Tag.ToString());
                            if (currentFrame < frameCount - 1)
                            {
                                currentFrame++;
                                imagePreview.Image = GetTiffFrame(tiffImage, currentFrame);
                                nextButton.Tag = currentFrame;
                            }
                        };

                        nextButton.Tag = 0; // Página inicial
                        multiPagePanel.Controls.Add(imagePreview);
                        multiPagePanel.Controls.Add(nextButton);
                        previewPanel.Controls.Add(multiPagePanel);
                    }
                    else
                    {
                        PictureBox imagePreview = new PictureBox
                        {
                            Dock = DockStyle.Fill,
                            SizeMode = PictureBoxSizeMode.Zoom,
                            Image = new Bitmap(tiffImage)
                        };
                        previewPanel.Controls.Add(imagePreview);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowUnsupportedMessage($"Error al cargar el archivo TIFF: {ex.Message}");
            }
        }

        // WIC - Windows Imaging Component
        // Variables globales para controlar el factor de zoom
        float zoomFactor = 1.0f;
        private void LoadTiffWithWIC(string fullPath, ToolStrip toolbar, Panel mainPanel)
        {
            try
            {
                // Crear el PictureBox para mostrar la imagen
                PictureBox imagePreview = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    //Dock = DockStyle.None, // Permite un control más preciso de las dimensiones
                };

                // Decodificar el archivo TIFF
                BitmapDecoder decoder = BitmapDecoder.Create(
                    new Uri(fullPath),
                    BitmapCreateOptions.PreservePixelFormat,
                    BitmapCacheOption.OnLoad);

                // Variables para navegación
                int currentPage = 0;
                int totalPages = decoder.Frames.Count;

                // Función para cargar una página específica
                void LoadPage(int pageIndex)
                {
                    // Convertir la primera página del TIFF a un Bitmap para el PictureBox
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        BitmapFrame frame = decoder.Frames[pageIndex]; // Obtener la primera página del TIFF
                        BitmapEncoder encoder = new BmpBitmapEncoder(); // Convertir a BMP
                        encoder.Frames.Add(frame);
                        encoder.Save(memoryStream);

                        // Crear un objeto Bitmap para el PictureBox
                        Bitmap bitmap = new Bitmap(memoryStream);

                        // Ajustar dinámicamente el tamaño del PictureBox para que la imagen quepa
                        AdjustImageToPanelSize(bitmap, imagePreview, mainPanel);

                        imagePreview.Image = bitmap; // Asignar al PictureBox
                        //imagePreview.Width = (int)(bitmap.Width * zoomFactor);
                        //imagePreview.Height = (int)(bitmap.Height * zoomFactor);
                    }

                    // Actualizar el estado de los botones
                    UpdateButtonState();
                }

                void AdjustImageToPanelSize(Bitmap bitmap, PictureBox pictureBox, Panel panel)
                {
                    // Obtener dimensiones del panel principal
                    int panelWidth = panel.ClientSize.Width;
                    int panelHeight = panel.ClientSize.Height;

                    // Calcular el factor de escala
                    float widthScale = (float)panelWidth / bitmap.Width;
                    float heightScale = (float)panelHeight / bitmap.Height;

                    // Usar el menor factor de escala para mantener la relación de aspecto
                    float scaleFactor = Math.Min(widthScale, heightScale);

                    // Ajustar el tamaño del PictureBox según el factor de escala
                    pictureBox.Width = (int)(bitmap.Width * scaleFactor);
                    pictureBox.Height = (int)(bitmap.Height * scaleFactor);

                    // Centrar el PictureBox dentro del panel
                    pictureBox.Left = (panelWidth - pictureBox.Width) / 2;
                    pictureBox.Top = (panelHeight - pictureBox.Height) / 2;
                }

                // Ruta a la carpeta de imágenes
                string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images");

                // Crear botón "Siguiente"
                ToolStripButton btnNext = new ToolStripButton
                {
                    Enabled = totalPages > 1,
                    Text = "",
                    DisplayStyle = ToolStripItemDisplayStyle.Image,
                    ToolTipText = "Página siguiente",
                    Image = Image.FromFile(Path.Combine(imagePath, "right-arrow.png")),
                    Margin = new Padding(5, 5, 5, 5),   // Espacio libre alrededor del botón
                    ImageScaling = ToolStripItemImageScaling.None,
                    AutoSize = false,
                    Width = 36,
                    Height = 36
                };

                // Asignar eventos para cambiar el cursor
                btnNext.MouseEnter += (s, e) => { Cursor.Current = Cursors.Hand; };  // Cambia a manoPrevious
                btnNext.MouseLeave += (s, e) => { Cursor.Current = Cursors.Default; };  // Restaura el cursor

                // Crear botón "Anterior"
                ToolStripButton btnPrevious = new ToolStripButton
                {
                    Enabled = false,    // Inicialmente desactivado,
                    Text = "",
                    DisplayStyle = ToolStripItemDisplayStyle.Image, // Solo mostrar la imagen
                    ToolTipText = "Página anterior",
                    Image = Image.FromFile(Path.Combine(imagePath, "back-arrow.png")),
                    Margin = new Padding(5, 5, 5, 5),   // Espacio libre alrededor del botón
                    ImageScaling = ToolStripItemImageScaling.None,
                    AutoSize = false,
                    Width = 36,
                    Height = 36,
                };

                // Asignar eventos para cambiar el cursor
                btnPrevious.MouseEnter += (s, e) => { Cursor.Current = Cursors.Hand; };  // Cambia a manoPrevious
                btnPrevious.MouseLeave += (s, e) => { Cursor.Current = Cursors.Default; };  // Restaura el cursor

                // Botón de imprimir
                ToolStripButton btnPrint = new ToolStripButton
                {
                    Text = "",
                    DisplayStyle = ToolStripItemDisplayStyle.Image,
                    ToolTipText = "Imprimir documento",
                    Image = Image.FromFile(Path.Combine(imagePath, "printer.png")),
                    Margin = new Padding(5, 5, 5, 5),   // Espacio libre alrededor del botón
                    ImageScaling = ToolStripItemImageScaling.None,
                    AutoSize = false,
                    Width = 36,
                    Height = 36
                };

                // Asignar eventos para cambiar el cursor
                btnPrint.MouseEnter += (s, e) => { Cursor.Current = Cursors.Hand; };  // Cambia a manoPrevious
                btnPrint.MouseLeave += (s, e) => { Cursor.Current = Cursors.Default; };  // Restaura el cursor

                // Agregar botones al panel
                toolbar.Items.Add(btnNext);
                toolbar.Items.Add(btnPrevious);
                toolbar.Items.Add(btnPrint);

                // Cargar la primera página
                LoadPage(currentPage);

                // NEXT
                btnNext.Click += (s, e) =>
                {
                    if (currentPage < totalPages - 1)
                    {
                        currentPage++;
                        LoadPage(currentPage);
                    }
                };

                // PREVIOUS
                btnPrevious.Click += (s, e) =>
                {
                    if (currentPage > 0)
                    {
                        currentPage--;
                        LoadPage(currentPage);
                    }
                };

                // PRINT
                btnPrint.Click += (s, e) =>
                {
                    PrintDocument printDocument = new PrintDocument();
                    printDocument.PrintPage += (sender, args) =>
                    {
                        args.Graphics.DrawImage(imagePreview.Image, args.MarginBounds);
                    };
                    PrintDialog printDialog = new PrintDialog
                    {
                        Document = printDocument
                    };
                    if (printDialog.ShowDialog() == DialogResult.OK)
                    {
                        printDocument.Print();
                    }
                };

                // Función para actualizar el estado de los botones
                void UpdateButtonState()
                {
                    btnPrevious.Enabled = currentPage > 0;
                    btnNext.Enabled = currentPage < totalPages - 1;
                }

                // Botón de zoom (aumentar)
                Button btnZoomIn = new Button
                {
                    Text = "Zoom +"
                };
                btnZoomIn.Click += (s, e) =>
                {
                    zoomFactor *= 1.2f; // Incrementa el factor de zoom
                    imagePreview.Dock = DockStyle.None;
                    imagePreview.Width = (int)(imagePreview.Width * zoomFactor);
                    imagePreview.Height = (int)(imagePreview.Height * zoomFactor);
                 };
                //toolbar.Controls.Add(btnZoomIn);

                // Botón de zoom (disminuir)
                Button btnZoomOut = new Button
                {
                    Text = "Zoom -"
                };
                btnZoomOut.Click += (s, e) =>
                {
                    zoomFactor /= 1.2f; // Decrementa el factor de zoom
                    imagePreview.Dock = DockStyle.None;
                    imagePreview.Width = (int)(imagePreview.Width * zoomFactor);
                    imagePreview.Height = (int)(imagePreview.Height * zoomFactor);
                };
                //toolbar.Controls.Add(btnZoomOut);

                // Agregar el PictureBox al mainPanel
                mainPanel.Controls.Add(imagePreview);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar el archivo TIFF con WIC: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Microsoft.Office.Interop.Word - renderizar documentos de Word
        private void ShowWordPreview(string fullPath, Panel mainPanel)
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Microsoft.Office.Interop.Word.Document wordDoc = null;
            string tempPdfPath = null;

            try
            {
                // Verificar si el archivo está bloqueado antes de abrirlo
                if (IsFileLocked(fullPath))
                {
                    MessageBox.Show("Документ використовується іншим процесом.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Crear una nueva instancia de Word
                wordApp = new Microsoft.Office.Interop.Word.Application();
                wordDoc = wordApp.Documents.Open(
                    fullPath,
                    ReadOnly: true, // Abrir en modo de solo lectura
                    Visible: false
                );

                /*
                // Exportar a PDF para previsualizar
                tempPdfPath = Path.ChangeExtension(fullPath, ".pdf");

                // Eliminar archivo temporal si existe y no está bloqueado
                if (!string.IsNullOrEmpty(tempPdfPath) && File.Exists(tempPdfPath))
                {
                    if (IsFileLocked(tempPdfPath))
                    {
                        // Si está bloqueado, esperar unos momentos o manejarlo de otra forma
                        MessageBox.Show("El archivo temporal PDF está bloqueado y no puede ser eliminado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        // return;
                    }

                    try
                    {
                        File.Delete(tempPdfPath);
                        // Sobrescribir el archivo directamente sin eliminarlo
                        // wordDoc.ExportAsFixedFormat(tempPdfPath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не можливо видалити тимчасовий документ: {ex.Message}", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                */

                // Generar un nombre único para el archivo temporal PDF
                string tempDirectory = Path.GetTempPath(); // Usar el directorio temporal del sistema
                tempPdfPath = Path.Combine(tempDirectory, $"{Guid.NewGuid()}.pdf");
                // Guardamos la reuta completa al fichero temporal para poder borarlo al final de la sesion
                tempFiles.Add(tempPdfPath);

                // Exportar el documento de Word a PDF
                wordDoc.ExportAsFixedFormat(tempPdfPath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                // Mostrar el PDF en un WebBrowser o cualquier visor de PDF
                WebBrowser pdfPreview = new WebBrowser
                {
                    Dock = DockStyle.Fill
                };
                pdfPreview.Navigate(tempPdfPath);
                mainPanel.Controls.Add(pdfPreview);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Помилка відображення файла .doc/.docx: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Liberar recursos
                try
                {
                    if (wordDoc != null)
                    {
                        wordDoc.Close(false);
                        Marshal.ReleaseComObject(wordDoc);
                    }
                    if (wordApp != null)
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Помилка звільнення ресурсів документа Word: {ex.Message}");
                }

                wordDoc = null;
                wordApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                /*
                // Eliminar archivo temporal
                if (!string.IsNullOrEmpty(tempPdfPath) && File.Exists(tempPdfPath))
                {
                    try
                    {
                        File.Delete(tempPdfPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Не можливо видалити тимчасовий документ: {ex.Message}");
                    }
                }
                */
            }
        }

        // Asegurarse si un archivo está bloqueado antes de intentar sobrescribirlo
        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    return false; // No está bloqueado
                }
            }
            catch (IOException)
            {
                return true; // Está bloqueado
            }
        }

        // Manejo adicional de procesos de Word
        private void KillExistingWordProcesses()
        {
            var wordProcesses = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (var process in wordProcesses)
            {
                try
                {
                    process.Kill();
                    process.WaitForExit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"No se pudo cerrar el proceso Word: {ex.Message}");
                }
            }
        }


        // Obtener un marco específico de un TIFF
        private Image GetTiffFrame(Image tiffImage, int frameIndex)
        {
            var dimension = new System.Drawing.Imaging.FrameDimension(tiffImage.FrameDimensionsList[0]);
            tiffImage.SelectActiveFrame(dimension, frameIndex);
            return new Bitmap(tiffImage); // Crear una copia en memoria del marco
        }

        // Mostrar un mensaje de error en caso de Vista previa del documento
        private void ShowUnsupportedMessage(string textError)
        {
            var label = new Label
            {
                Text = textError,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            previewPanel.Controls.Add(label);
        }

        // Sobrescribir OnShown
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            // Actualizar la duración del inicio de la aplicación
            _mainApp.StartAppEnd();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Lógica de cierre sin confirmación, ya que esta se maneja en OnFormClosing
            _mainApp.EndApp(); // Cierra la sesión actual de SQL Server y libera recursos

            Application.Exit(); // Cierra la aplicación

        }

        // Usar el evento Paint del formulario para dibujar un rectángulo sobre el área del divisor
        // para colorar el Split en color Blue
        private void splitContainer1_Paint(object sender, PaintEventArgs e)
        {
            // Obtener las dimensiones del divisor
            int splitterWidth = splitContainer1.SplitterWidth;
            int splitterPosition = splitContainer1.SplitterDistance;

            // Crear un pincel con el color deseado
            using (Brush brush = new SolidBrush(Color.Blue)) // Cambia el color aquí
            {
                // Dibujar un rectángulo sobre el divisor
                if (splitContainer1.Orientation == Orientation.Vertical)
                {
                    e.Graphics.FillRectangle(brush, splitterPosition, 0, splitterWidth, splitContainer1.Height);
                }
                else
                {
                    e.Graphics.FillRectangle(brush, 0, splitterPosition, splitContainer1.Width, splitterWidth);
                }
            }
        }

        // Configurar la vista de treeView para que se vea más agradable y moderno
        private void ConfigureTreeView()
        {
            // Configurar las propiedades básicas del TreeView
            treeView.ItemHeight = 36; // Altura de los elementos
            treeView.ShowLines = true; // Mostrar líneas entre nodos
            treeView.HideSelection = false; // Mostrar selección incluso cuando el TreeView pierde el foco
            treeView.BackColor = Color.WhiteSmoke; // Fondo moderno
            treeView.ForeColor = Color.DarkSlateGray; // Color del texto

            // Estilo de fuente ajustado para un aspecto más limpio
            treeView.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            // Configuración del ImageList (asumiendo que ya está inicializado)
            if (_imageList != null)
            {
                _imageList.ImageSize = new Size(32, 32); // Asegúrate de que las imágenes tengan el tamaño correcto
                treeView.ImageList = _imageList;
            }

            // Configuración adicional si es necesario
        }

        // Ejecuta cuando se pulsa la cruz de cerrar la app
        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            // Verificar si ya se está cerrando la aplicación
            if (_isClosing)
            {
                return;
            }

            // Preguntar si el usuario quiere cerrar la aplicación
            var result = MessageBox.Show(
                "Ви впевнені, що бажаєте вийти з програим?",
                "Підтвердження виходу",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Borramos los ficheros temporales
                foreach (var tempFile in tempFiles)
                {
                    if (File.Exists(tempFile))
                    {
                        try
                        {
                            File.Delete(tempFile);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"No se pudo eliminar el archivo temporal {tempFile}: {ex.Message}");
                        }
                    }
                }

                // Establece la bandera para evitar múltiples ejecuciones
                _isClosing = true;

                // Llama a la función exitToolStripMenuItem_Click para liberar recursos
                exitToolStripMenuItem_Click(null, EventArgs.Empty);

                // Permitir el cierre del formulario
                e.Cancel = false;
            }
            else
            {
                // Cancela el cierre del formulario
                e.Cancel = true;
            }
        }
    }
}
