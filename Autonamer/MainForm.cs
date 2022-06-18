/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 22/02/2016
 * Hora: 09:53 a. m.
 * 
 * Para cambiar esta plantilla use Herramientas | Opciones | Codificación | Editar Encabezados Estándar
 */
using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Threading;
using System.Diagnostics;


namespace Autonamer
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		String nombre_viejo,nombre_nuevo,ruta,extension,ext,arch_ex,cad_con,tabla,hoja,cons_exc;
        int i=0,j=0,filas=0,tot_rows=0,camb_ext=0;
        string[] archivos_tot;
        string[] por_borrar;
        
        //caracteres prohibidos en los nombres: 
        // :,\,/,*,?,",<,>,|
        
        
        //Declaracion de elementos para conexion office
		OleDbConnection conexion = null;
		DataSet dataSet = null;
		OleDbDataAdapter dataAdapter = null;
        
		public void abrir_carpeta(){
			i=0;
			dataGridView1.Rows.Clear();
        	FolderBrowserDialog fbd = new FolderBrowserDialog();
        	fbd.Description = "Selecciona la carpeta que contiene los archivos a renombrar";
        	fbd.ShowNewFolderButton =false;
        	DialogResult result = fbd.ShowDialog();
        	
        	if(result == DialogResult.OK){
        		archivos_tot = Directory.GetFiles(fbd.SelectedPath);
        		//MessageBox.Show("Archivo 1: "+archivos_tot[0]);
        		
        		button2.Enabled=true;
        		checkBox1.Enabled=true;
        		button3.Enabled=true;
        		label2.Text="Archivos Cargados: "+archivos_tot.Length;
        		
        		do{
        			dataGridView1.Rows.Add();
        			nombre_viejo = archivos_tot[i];
        			ruta = nombre_viejo.Substring(0,(nombre_viejo.LastIndexOf('\\')+1));
        			if((nombre_viejo.LastIndexOf('.') > -1)){
        				extension = nombre_viejo.Substring((nombre_viejo.LastIndexOf('.')+1),(nombre_viejo.Length-(nombre_viejo.LastIndexOf('.')+1)));
        				nombre_viejo=nombre_viejo.Substring((nombre_viejo.LastIndexOf('\\')+1),((nombre_viejo.LastIndexOf('.'))-nombre_viejo.LastIndexOf('\\'))-1);
        			}else{
        			   	extension = " ";
        			   	nombre_viejo=nombre_viejo.Substring((nombre_viejo.LastIndexOf('\\')+1),((nombre_viejo.Length)-nombre_viejo.LastIndexOf('\\'))-1);
        			}
        			
        			
        			dataGridView1.Rows[i].Cells[0].Value=nombre_viejo;
        			dataGridView1.Rows[i].Cells[2].Value=archivos_tot[i];
        			
        			i++;
        		}while(i<archivos_tot.Length);
        		dataGridView1.Sort(Column1,System.ComponentModel.ListSortDirection.Ascending);
        		
        		//MessageBox.Show(extension);
        	}else{
        		MessageBox.Show("No se hará nada   :(");
        		
        	}
		}
        
        public void carga_excel(){
			OpenFileDialog dialog = new OpenFileDialog();
			dialog.Filter = "Archivos de Excel (*.xls *.xlsx)|*.xls;*.xlsx"; //le indicamos el tipo de filtro en este caso que busque
			//solo los archivos excel
			dialog.Title = "Seleccione el archivo de Excel";//le damos un titulo a la ventana
			dialog.FileName = string.Empty;//inicializamos con vacio el nombre del archivo
			
			//si al seleccionar el archivo damos Ok
			if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				arch_ex = dialog.FileName;
				//label2.Text = dialog.SafeFileName;
				ext=arch_ex.Substring(((arch_ex.Length)-3),3);
				ext=ext.ToLower();
				
				if(ext.Equals("lsx")){
					MessageBox.Show("Asegurate de Cerrar el archivo en Excel, Antes de abrirlo aqui","Advertencia");
				}
				
				//esta cadena es para archivos excel 2007 y 2010
				cad_con = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + arch_ex + "';Extended Properties=Excel 12.0;";
				conexion = new OleDbConnection(cad_con);//creamos la conexion con la hoja de excel
				conexion.Open(); //abrimos la conexion
				
				carga_chema_excel();
				
			}
		}
		
		public void carga_chema_excel(){
			i=0;
			filas = 0;
			comboBox1.Items.Clear();
			System.Data.DataTable dt = conexion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
		    dataGridView2.DataSource =dt;
		    filas=(dataGridView2.RowCount)-1;
					do{
						if (!(dataGridView2.Rows[i].Cells[3].Value.ToString()).Equals("")){
							if ((dataGridView2.Rows[i].Cells[3].Value.ToString()).Equals("TABLE")){
								tabla=dataGridView2.Rows[i].Cells[2].Value.ToString();
								if((tabla.Substring((tabla.Length-1),1)).Equals("$")){
									tabla = tabla.Remove((tabla.Length-1),1);
									comboBox1.Items.Add(tabla);
								}
							}
						}
						i++;
					}while(i<=filas);
					
                    //dt.Clear();
                    //dataGridView2.DataSource = dt; //vaciar datagrid
                    comboBox1.Enabled=true;
		}
		
		public void cargar_hoja_excel(){
			
				hoja = comboBox1.SelectedItem.ToString();
			
			if (string.IsNullOrEmpty(hoja))
			{
				MessageBox.Show("No hay una hoja para leer");
			}
			else
			{
				cons_exc = "Select * from [" + hoja + "$] ";
				
				try
				{
					//Si el usuario escribio el nombre de la hoja se procedera con la busqueda
					//conexion = new OleDbConnection(cadenaConexionArchivoExcel);//creamos la conexion con la hoja de excel
					//conexion.Open(); //abrimos la conexion
					dataAdapter = new OleDbDataAdapter(cons_exc, conexion); //traemos los datos de la hoja y las guardamos en un dataSdapter
					dataSet = new DataSet(); // creamos la instancia del objeto DataSet
					if(dataAdapter.Equals(null)){
						
						MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja\n","Error al Abrir Archivo de Excel/");
						
					}else{
						dataAdapter.Fill(dataSet, hoja);//llenamos el dataset
						dataGridView3.DataSource = dataSet.Tables[0]; //le asignamos al DataGridView el contenido del dataSet
						conexion.Close();//cerramos la conexion
						dataGridView1.AllowUserToAddRows = false;       //eliminamos la ultima fila del datagridview que se autoagrega
						tot_rows=dataGridView3.RowCount;
						//label2.Text="Registros: "+tot_rows;
						//label2.Refresh();
						

						//estilo datagrid
                       
                       

						if(tot_rows>0){
							//maskedTextBox2.Enabled=true;
							//comboBox3.Enabled=true;							
						}else{
							//maskedTextBox2.Enabled=false;
							//comboBox3.Enabled=false;
							//button10.Enabled=false;
							//button4.Enabled=false;
						}                        
					}
				}
				catch (AccessViolationException ex )
				{
					//en caso de haber una excepcion que nos mande un mensaje de error
					MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja\n"+ex,"Error al Abrir Archivo de Excel");
				}
				
			}
			
		}
		
		public void cargar_hoja_en_grid(){
			i=0;
			do{
				if(i<dataGridView3.RowCount){
					nombre_viejo = dataGridView3.Rows[i].Cells[0].Value.ToString();
				}else{
					nombre_viejo = " ";
				}
        			//ruta = nombre_viejo.Substring(0,(nombre_viejo.LastIndexOf('\\')+1));
        			//extension= nombre_viejo.Substring((nombre_viejo.LastIndexOf('.')+1),(nombre_viejo.Length-(nombre_viejo.LastIndexOf('.')+1)));
        			//nombre_viejo=nombre_viejo.Substring((nombre_viejo.LastIndexOf('\\')+1),((nombre_viejo.LastIndexOf('.'))-nombre_viejo.LastIndexOf('\\'))-1);
        			
        			dataGridView1.Rows[i].Cells[1].Value=nombre_viejo;
        			i++;
        		}while(i<dataGridView1.RowCount);
		}
        
		public void aviso(){
			
			DialogResult respuesta;
			
			if(camb_ext==0){
				respuesta = MessageBox.Show("Estás apunto de cambiar el nombre de "+dataGridView1.RowCount+" archivos.\n"+
				                            "\n¿Deseas continuar?","CONFIRMAR",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
					
				if(respuesta == DialogResult.Yes){
					//MessageBox.Show("lalalala :(");
					rename();
				}else{
					MessageBox.Show("No se hizo nada  :(");
				}
				
			}else{
				respuesta = MessageBox.Show("Estás apunto de cambiar el nombre y la extension de "+dataGridView1.RowCount+" archivos.\n"+
				                            "Al Modificar la extensión de un archivo se arriesga a dejarlo ilegible.\n"+
				                            "\n¿Deseas continuar?","CONFIRMAR",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
					
				if(respuesta == DialogResult.Yes){
					rename_con_ext();
				}else{
					MessageBox.Show("No se hizo nada  :(");
				}
			}
		}
		
		public void rename(){
			i=0;
			j=0;
			por_borrar = new string[dataGridView1.RowCount];
			
			do{
				if(dataGridView1.Rows[i].Cells[1].Value != null){
					nombre_nuevo = dataGridView1.Rows[i].Cells[1].Value.ToString();
					//MessageBox.Show(nombre_viejo+","+nombre_nuevo);
					if((nombre_nuevo.Contains("\\"))||(nombre_nuevo.Contains("/"))||(nombre_nuevo.Contains("*"))||(nombre_nuevo.Contains("?"))){
						
					}else{
                        if ((nombre_nuevo.Contains("\"")) || (nombre_nuevo.Contains("<")) || (nombre_nuevo.Contains(">")) || (nombre_nuevo.Contains("|")) || (nombre_nuevo.Contains(":")))
                        {

						}else{
							nombre_viejo = dataGridView1.Rows[i].Cells[2].Value.ToString();
							ruta = nombre_viejo.Substring(0,(nombre_viejo.LastIndexOf('\\')+1));
							
							if((nombre_viejo.LastIndexOf('.') > -1)){
								extension = nombre_viejo.Substring((nombre_viejo.LastIndexOf('.')+1),(nombre_viejo.Length-(nombre_viejo.LastIndexOf('.')+1)));
								nombre_nuevo=ruta+nombre_nuevo+"."+extension;
							}else{
								extension = " ";
								nombre_nuevo=ruta+nombre_nuevo;
							}
	
							//MessageBox.Show(nombre_viejo+","+nombre_nuevo);
							if((System.IO.File.Exists(nombre_nuevo))==false){
								System.IO.File.Move(nombre_viejo,nombre_nuevo);		
		        				por_borrar[j]=i.ToString();
								j++;
							}
						}
					}
				}else{
					
				}
				
				i++;
			}while(i<dataGridView1.RowCount);
			
			if(j>0){
				i=j-1;
				do{
					dataGridView1.Rows.RemoveAt(Convert.ToInt32(por_borrar[i]));
					i--;
				}while(i>-1);
			}
			
			MessageBox.Show("El proceso ha terminado adecuadamente.\n"+
			                "Se Renombraron correctamente: "+j+" archivos.\n"+
			                "Se Omitieron: "+dataGridView1.RowCount+" archivos por no contar con un nuevo nombre válido","Listo!");
		
		    label2.Text="Archivos Cargados: "+dataGridView1.RowCount;
		    if(dataGridView1.RowCount==0){
				button2.Enabled=false;
        		checkBox1.Enabled=false;
        		button3.Enabled=false;
        		comboBox1.Enabled=false;
        		checkBox1.Checked=false;
        		comboBox1.Items.Clear();
			}
		}
		
		public void rename_con_ext(){
			i=0;
			j=0;
			por_borrar = new string[dataGridView1.RowCount];
			
			do{
				if(dataGridView1.Rows[i].Cells[1].Value != null){
					nombre_nuevo = dataGridView1.Rows[i].Cells[1].Value.ToString();
					//MessageBox.Show(nombre_viejo+","+nombre_nuevo);
					if((nombre_nuevo.Contains("\\"))||(nombre_nuevo.Contains("/"))||(nombre_nuevo.Contains("*"))||(nombre_nuevo.Contains("?"))){
						
					}else{
                        if ((nombre_nuevo.Contains("\"")) || (nombre_nuevo.Contains("<")) || (nombre_nuevo.Contains(">")) || (nombre_nuevo.Contains("|")) || (nombre_nuevo.Contains(":")))
                        {

						}else{
							nombre_viejo = dataGridView1.Rows[i].Cells[2].Value.ToString();
							ruta = nombre_viejo.Substring(0,(nombre_viejo.LastIndexOf('\\')+1));
							
							nombre_nuevo=ruta+nombre_nuevo;
							
							//MessageBox.Show(nombre_viejo+","+nombre_nuevo);
							if((System.IO.File.Exists(nombre_nuevo))==false){
								System.IO.File.Move(nombre_viejo,nombre_nuevo);		
		        				por_borrar[j]=i.ToString();
								j++;
							}
						}
					}
				}else{
					
				}
				
				i++;
			}while(i<dataGridView1.RowCount);
			
			if(j>0){
				i=j-1;
				do{
					dataGridView1.Rows.RemoveAt(Convert.ToInt32(por_borrar[i]));
					i--;
				}while(i>-1);
			}
			
			MessageBox.Show("El proceso ha terminado adecuadamente.\n"+
			                "Se Renombraron correctamente: "+j+" archivos.\n"+
			                "Se Omitieron: "+dataGridView1.RowCount+" archivos por no contar con un nuevo nombre válido","Listo!");
		
		    label2.Text="Archivos Cargados: "+dataGridView1.RowCount;
		    if(dataGridView1.RowCount==0){
				button2.Enabled=false;
        		checkBox1.Enabled=false;
        		button3.Enabled=false;
        		comboBox1.Enabled=false;
        		checkBox1.Checked=false;
        		comboBox1.Items.Clear();
			}
		}
		
		void Button1Click(object sender, EventArgs e)
		{
			abrir_carpeta();
		}
		
		void MainFormLoad(object sender, EventArgs e)
		{
			
		}
		
		void Button2Click(object sender, EventArgs e)
		{
			carga_excel();
		}
		
		void Button3Click(object sender, EventArgs e)
		{
			aviso();
		}
		
		void CheckBox1CheckedChanged(object sender, EventArgs e)
		{
			if(dataGridView1.RowCount > 0){
				i=0;
				if(checkBox1.Checked==true){
					do{
						nombre_viejo = dataGridView1.Rows[i].Cells[2].Value.ToString();
						ruta = nombre_viejo.Substring(0,(nombre_viejo.LastIndexOf('\\')+1));
						
						if((nombre_viejo.LastIndexOf('.') > -1)){
							extension = nombre_viejo.Substring((nombre_viejo.LastIndexOf('.')+1),(nombre_viejo.Length-(nombre_viejo.LastIndexOf('.')+1)));
							nombre_viejo=nombre_viejo.Substring((nombre_viejo.LastIndexOf('\\')+1),((nombre_viejo.LastIndexOf('.'))-nombre_viejo.LastIndexOf('\\'))-1);
						}else{
							extension = " ";
							nombre_viejo=nombre_viejo.Substring((nombre_viejo.LastIndexOf('\\')+1),((nombre_viejo.Length)-nombre_viejo.LastIndexOf('\\'))-1);
						}
						
						dataGridView1.Rows[i].Cells[0].Value=nombre_viejo+"."+extension;
						i++;
					}while(i<dataGridView1.RowCount);
					camb_ext=1;
				}else{
					do{
						nombre_viejo = dataGridView1.Rows[i].Cells[2].Value.ToString();
						ruta = nombre_viejo.Substring(0,(nombre_viejo.LastIndexOf('\\')+1));
						
						if((nombre_viejo.LastIndexOf('.') > -1)){
							extension = nombre_viejo.Substring((nombre_viejo.LastIndexOf('.')+1),(nombre_viejo.Length-(nombre_viejo.LastIndexOf('.')+1)));
							nombre_viejo=nombre_viejo.Substring((nombre_viejo.LastIndexOf('\\')+1),((nombre_viejo.LastIndexOf('.'))-nombre_viejo.LastIndexOf('\\'))-1);
						}else{
							extension = " ";
							nombre_viejo=nombre_viejo.Substring((nombre_viejo.LastIndexOf('\\')+1),((nombre_viejo.Length)-nombre_viejo.LastIndexOf('\\'))-1);
						}
						
						dataGridView1.Rows[i].Cells[0].Value=nombre_viejo;
						i++;
					}while(i<dataGridView1.RowCount);
					camb_ext=0;
				}
				dataGridView1.Sort(Column1,System.ComponentModel.ListSortDirection.Ascending);
			}
		}
		
		void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
		{
			cargar_hoja_excel();
			cargar_hoja_en_grid();
		}
		
		void DataGridView1UserDeletedRow(object sender, DataGridViewRowEventArgs e)
		{
			label2.Text="Archivos Cargados: "+dataGridView1.RowCount;
			if(dataGridView1.RowCount==0){
				button2.Enabled=false;
        		checkBox1.Enabled=false;
        		button3.Enabled=false;
        		comboBox1.Enabled=false;
        		checkBox1.Checked=false;
        		comboBox1.Items.Clear();
			}
		}
		
		void Button4Click(object sender, EventArgs e)
		{
			About abo = new About();
			abo.ShowDialog();
		}
		
		void DataGridView1CellClick(object sender, DataGridViewCellEventArgs e)
		{
			
		}
		
		void CopiaToolStripMenuItemClick(object sender, EventArgs e)
		{
			if(dataGridView1.RowCount>0){
				dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
				label2.Text="Archivos Cargados: "+dataGridView1.RowCount;
			}
			
			if(dataGridView1.RowCount==0){
				button2.Enabled=false;
        		checkBox1.Enabled=false;
        		button3.Enabled=false;
        		comboBox1.Enabled=false;
        		checkBox1.Checked=false;
        		comboBox1.Items.Clear();
			}
		}
		
		void DataGridView1RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
		{
			dataGridView1.CurrentRow.Selected=true;
		}
		
		void DataGridView1CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
		{
			
		}
		
		void DataGridView1MouseClick(object sender, MouseEventArgs e)
		{
			
			
		}
		
		void DataGridView1MouseDown(object sender, MouseEventArgs e)
		{
			if(dataGridView1.RowCount>0){
				if(e.Button == MouseButtons.Right){
                    dataGridView1.MultiSelect = false;
                    dataGridView1.MultiSelect = true;
                    if ((0 <= (dataGridView1.HitTest(e.X, e.Y).RowIndex)) && ((dataGridView1.HitTest(e.X, e.Y).RowIndex) <= (dataGridView1.RowCount-1)))
                    {
					dataGridView1.Rows[dataGridView1.HitTest(e.X,e.Y).RowIndex].Selected=true;
                    }
					//MessageBox.Show("Fila numero: "+(dataGridView1.HitTest(e.X,e.Y).RowIndex+1));
				}
			}
		}

        private void button5_Click(object sender, EventArgs e)
        {
            Instrucciones ins = new Instrucciones();
            ins.Show();
        }
	}
}
