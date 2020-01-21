package com.uds.detailForm.handlers;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.WindowConstants;
import javax.swing.table.DefaultTableModel;

import ngc_utils.Common;
import ngc_utils.JDBCUtils;
import ngc_utils.ReportCommon;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.TCTypeService;
import com.teamcenter.rac.util.MessageBox;

/**
 * ������ϸ��
 * @author Guoziang
 *
 * 2018-11-28
 */
public class SparePartsForGY extends AbstractHandler{
	
	AbstractAIFUIApplication app = null;
	TCSession session = null;
	String[] InFileName = null;
	String TempPath = "c:\\temp\\";
	Workbook wb = null;
	// ������ϸ��
	Sheet sheet1 = null;
	// ��ƷͼֽĿ¼
	Sheet sheet2 = null;
	// �汾�����¼
	Sheet sheet3 = null;
	// ����
	Sheet sheet4 = null;
	InterfaceAIFComponent selComp = null;
	TCComponentItemRevision tcItemrev = null;
	String OutFileName = "";
	private ArrayList<TCComponentItemRevision> relatedChangeComponentItemRevisionList;
	//private List<RevBean> revLineList = new ArrayList<>();
	TCComponentBOMLine selectedBOMLine = null;
	String whole_nc8_drawing_no = "";
	private String nc8_order_line_number = "";
	private String temp_nc8_order_number = "";
	private String sessionUserName = "";
	static int rowNum = 4;
	static int number = 1;
	static int productRowNum = 5;
	private String orderNo="";
	private String userStr = "";
	
//	List<TCComponentBOMLine> childBom = null;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {

		rowNum = 4;
		number = 1;
		productRowNum = 5;

		app = AIFDesktop.getActiveDesktop().getCurrentApplication();
		session = (TCSession) app.getSession();
		sessionUserName = session.getUser().toString();
		userStr = sessionUserName.substring(0,sessionUserName.indexOf("("));
		System.out.println(userStr+"===========");
		try {
			
			initUI();

		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	
	private void initUI() throws TCException{  
	       final JFrame jf = new JFrame("�����Ų�ѯ");
	       jf.setSize(500, 200);
	       jf.setResizable(false);
	       jf.setLocationRelativeTo(null);
	       
	       JPanel panel = new JPanel();
	       panel.setLayout(null);
	      
	       Integer START_X=130;
	       final JLabel jl = new JLabel("�����붩����");
	       jl.setBounds(START_X, 20, 230, 30);
	       
	       final JTextField tf=new JTextField();
	       tf.setBounds(START_X, 50, 230, 30);
	       
	       JButton btnOK = new JButton("��ѯ");
	       btnOK.addActionListener(new ActionListener() {    
				public void actionPerformed(ActionEvent e) {
					orderNo=new StringBuffer().append(tf.getText()).toString();
					if (orderNo.length()<1) {
						JOptionPane.showMessageDialog(null, "�����붩���ź�����!");
						return;
					}
					
					
					//	ͨ�������ŵ���Packing List...��ѯ����ѯ����װ�䵥
					TCComponent[] packings = Common.CommonFinder("Packing List...", "nc8_Sales_Order_No", orderNo);
					if (null != packings) {
						TCComponentItemRevision packing = (TCComponentItemRevision) packings[0];
						try {
							String nc8_Line_No = packing.getProperty("nc8_Line_No");
							String[] orderLineNums = nc8_Line_No.split(",");
							
							//	��orderLineNums��С�����Ÿ���
							if (null != orderLineNums) {
								for(int i = 0; i<orderLineNums.length - 1; i++) {
									for (int j = i+1; j < orderLineNums.length; j++) {
										if(Integer.parseInt(orderLineNums[i].trim()) > Integer.parseInt(orderLineNums[j].trim())) {
											String temp = orderLineNums[i];
											orderLineNums[i] = orderLineNums[j];
											orderLineNums[j] = temp;
										}
									}
								}
							}
							
							ArrayList<SparePartsInfoBean> aL = new ArrayList<SparePartsInfoBean>();
							
							for (int j = 0; j < orderLineNums.length; j++) {
								//	ͨ�������źͶ����кŵ���007-������ѯ������
								TCComponent[] topComponent = Common.CommonFinder("007-����", "nc8_order_number,nc8_order_line_number", orderNo + "," + orderLineNums[j]);
								if (null != topComponent) {
									TCComponentItemRevision orderItemRev = (TCComponentItemRevision) topComponent[0];
									String nc8_material_code = orderItemRev.getProperty("nc8_material_code");
									String nc8_model_no1 = orderItemRev.getProperty("nc8_model_no1");
									
									SparePartsInfoBean apib = new SparePartsInfoBean();
									apib.setNc8_order_number(orderNo);
									apib.setNc8_order_line_number(orderLineNums[j]);
									apib.setNc8_material_code(nc8_material_code);
									apib.setNc8_model_no1(nc8_model_no1);
									aL.add(apib);
									
								}
								
							}
							
							// ��ѡ��Ĵ���,���û�ѡ��Ҫ������Щ��
							querySparePartsUI(aL);
							
						} catch (TCException e1) {
							e1.printStackTrace();
						}
					}else{
							MessageBox.post("δ�ҵ�[������]Ϊ"+orderNo+"�Ķ�������!", "��ʾ",MessageBox.INFORMATION);
							jf.dispose();
							tf.setText("");
						}
					
					jf.dispose();
					
				}

	       });
	       btnOK.setBounds(300,120,80,30);
	       
	       JButton btnCancel = new JButton("ȡ��");
	       btnCancel.addActionListener(new ActionListener() {    
				public void actionPerformed(ActionEvent e) {
					jf.dispose();
				}
	       });
	       btnCancel.setBounds(120,120,80,30);
	       
	       panel.add(tf);
	       panel.add(jl);
	       panel.add(btnOK);
	       panel.add(btnCancel);
	       jf.setContentPane(panel);
	       jf.setVisible(true);
		}
	
	private void traverseBom(TCComponentBOMLine topBomline,int Level,TCComponentItemRevision orderItemRev, List<RevBean> bomLineListTest) throws TCException{
		
		String sequenceNo = null;//Bom���� �㼶
		TCComponentBOMLine parentBomLine = null;//Bom ��Bom
		String parentMaterialCode = null;//Bom���� �������ϱ���
		String bomQuantity = null;//Bom���� ����
		String bomRemark = null;//Bom���� ��ע
		String woodPattern = null;//Bom���� ľģͼ��
		
		//Bom���� �㼶
		sequenceNo = topBomline.getProperty("bl_sequence_no");
		
		//Bom���� �������ϱ���
		parentBomLine = topBomline.parent();
		if (parentBomLine != null) {
			TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
			if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
				parentMaterialCode = parentItemRevision.getProperty("nc8_material_code");
			} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
				parentMaterialCode = parentItemRevision.getProperty("nc8_material_number");
			}
		} else {
			System.out.println("��û���ϲ��������");
			parentMaterialCode = "";
		}
		
		//Bom���� ����
		bomQuantity = topBomline.getProperty("bl_quantity");
		if (bomQuantity == null || bomQuantity.length() == 0) {
			bomQuantity = "1";
		}
		
		//Bom���� ��ע
		bomRemark = topBomline.getProperty("NC8_BOM_remark");
		
		//Bom���� ľģͼ��
		if (topBomline.isValidPropertyName("nc8_wood_pattern") && (topBomline.getProperty("nc8_wood_pattern").length() != 0)) {
			woodPattern = topBomline.getProperty("nc8_wood_pattern");
		}
		
		
		RevBean revlinestruct = new RevBean(topBomline.getItemRevision(), Level,orderItemRev,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
		Add2BOMViewList(revlinestruct, bomLineListTest);
		
		AIFComponentContext[] aifComponent = topBomline.getChildren();
		if(aifComponent != null){
			for (int m = 0; m < aifComponent.length; m++) {
				TCComponentBOMLine childBomline = (TCComponentBOMLine) aifComponent[m].getComponent();
				if (!childBomline.isRoot()) {
					// �ж��Ƿ�չ��
					String NC8_autoExpand_true = childBomline.getProperty("NC8_autoExpand_true");
					if (!"��".equals(NC8_autoExpand_true)) {
//						RevBean revlinestruct = new RevBean(childBomline.getItemRevision(), Level+1,orderItemRev,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
//						Add2BOMViewList(revlinestruct, bomLineListTest);
						traverseBom(childBomline, Level + 1,orderItemRev, bomLineListTest);
					} else {
						
						//Bom���� �㼶
						sequenceNo = childBomline.getProperty("bl_sequence_no");
						//Bom���� �������ϱ���
						parentBomLine = childBomline.parent();
						if (parentBomLine != null) {
							TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
							if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
								parentMaterialCode = parentItemRevision.getProperty("nc8_material_code");
							} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
								parentMaterialCode = parentItemRevision.getProperty("nc8_material_number");
							}
						} else {
							System.out.println("��û���ϲ��������");
							parentMaterialCode = "";
						}
						
						//Bom���� ����
						bomQuantity = childBomline.getProperty("bl_quantity");
						if (bomQuantity == null || bomQuantity.length() == 0) {
							bomQuantity = "1";
						}
						
						//Bom���� ��ע
						bomRemark = childBomline.getProperty("NC8_BOM_remark");
						
						//Bom���� ľģͼ��
						if (childBomline.isValidPropertyName("nc8_wood_pattern") && (childBomline.getProperty("nc8_wood_pattern").length() != 0)) {
							woodPattern = childBomline.getProperty("nc8_wood_pattern");
						}
						
						RevBean childRevlinestruct = new RevBean(childBomline.getItemRevision(), Level+1,orderItemRev,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
						Add2BOMViewList(childRevlinestruct, bomLineListTest);
						checkYOrNExpand(childBomline, Level + 1, orderItemRev, bomLineListTest);
					}
				}
				
			}
		}
		
	}
	
	void Add2BOMViewList(RevBean newBOMLine, List<RevBean> bomLineList) {
//		boolean HasFind = false;
//		for (int i = 0; i < bomLineList.size(); i++) {
//			RevBean oldBOMLine = (RevBean) bomLineList.get(i);
//			if (oldBOMLine.RevLine == newBOMLine.RevLine) {
//				if (oldBOMLine.Level < newBOMLine.Level) {
//					oldBOMLine.Level = newBOMLine.Level;
//				}
//				HasFind = true;
//				break;
//			}
//		}
//		if (!HasFind) {
			bomLineList.add(newBOMLine);
//		}
	}
	
	public void checkYOrNExpand(TCComponentBOMLine BOMLine,int Level, TCComponentItemRevision orderItemRev ,List<RevBean> bomLineListTest) throws TCException {
		AIFComponentContext[] Cmp = BOMLine.getChildren();
		for (int i = 0; i < Cmp.length; i++) {
			TCComponentBOMLine childBOMLine = (TCComponentBOMLine) Cmp[i].getComponent();
			// �ж��Ƿ�չ��
			String NC8_autoExpand_true = childBOMLine.getProperty("NC8_autoExpand_true");
//			if ("��".equals(NC8_autoExpand_true)) {
				//�жϡ���չ���������е�ֵΪ���ǡ�
				String NC8_Y_or_N_Expand = childBOMLine.getProperty("NC8_Y_or_N_Expand");
				if ("��".equals(NC8_Y_or_N_Expand)) {
					//traverseBom(childBOMLine, Level + 1,orderItemRev, bomLineListTest);
					
					
					String sequenceNo = null;//Bom���� �㼶
					TCComponentBOMLine parentBomLine = null;//Bom ��Bom
					String parentMaterialCode = null;//Bom���� �������ϱ���
					String bomQuantity = null;//Bom���� ����
					String bomRemark = null;//Bom���� ��ע
					String woodPattern = null;//Bom���� ľģͼ��
					
					//Bom���� �㼶
					sequenceNo = childBOMLine.getProperty("bl_sequence_no");
					//Bom���� �������ϱ���
					parentBomLine = childBOMLine.parent();
					if (parentBomLine != null) {
						TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
						if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
							parentMaterialCode = parentItemRevision.getProperty("nc8_material_code");
						} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
							parentMaterialCode = parentItemRevision.getProperty("nc8_material_number");
						}
					} else {
						System.out.println("��û���ϲ��������");
						parentMaterialCode = "";
					}
					//Bom���� ����
					bomQuantity = childBOMLine.getProperty("bl_quantity");
					if (bomQuantity == null || bomQuantity.length() == 0) {
						bomQuantity = "1";
					}
					//Bom���� ��ע
					bomRemark = childBOMLine.getProperty("NC8_BOM_remark");
					//Bom���� ľģͼ��
					if (childBOMLine.isValidPropertyName("nc8_wood_pattern") && (childBOMLine.getProperty("nc8_wood_pattern").length() != 0)) {
						woodPattern = childBOMLine.getProperty("nc8_wood_pattern");
					}
					RevBean childRevlinestruct = new RevBean(childBOMLine.getItemRevision(), Level+1,orderItemRev,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
					Add2BOMViewList(childRevlinestruct, bomLineListTest);
				}
//			}
			checkYOrNExpand(childBOMLine,Level + 1,orderItemRev, bomLineListTest);
		}
		
	}
	
	private void writeDataToExcel(List<RevBean> revLineListTest) throws TCException, InvalidFormatException, IOException {
		
//		ColletcBOMView(selectedBOMLine, 1, bomLineListTest);

		// Collections.sort(bomLineList, new SortDataset());

		System.out.println("��revLineListTest��start");
		for (int i = 0; i < revLineListTest.size(); i++) {
			RevBean revStruct = revLineListTest.get(i);
			TCComponentItemRevision revLine = revStruct.RevLine;
//			TCComponentItemRevision bomLineItemRevision = bomLine.getItemRevision();
//			boolean idHasStatus = idHasStatus(bomLineItemRevision);
//			if (idHasStatus) {
//				MessageBox.post("BOM�ṹ���з������ϣ�����", "����", 1);
//				return;
//			}
			String name = revLine.getProperty("object_name");
			System.out.println("��index = " + i + ", name = " + name + ", level = " + revStruct.Level + "��");

		}
		System.out.println("��revLineListTest��end");

//		Shell shell = new Shell();
//		org.eclipse.swt.widgets.MessageBox messageBox = new org.eclipse.swt.widgets.MessageBox(shell, SWT.OK | SWT.CANCEL);
//		messageBox.setText("��ʾ");
//		messageBox.setMessage("�Ƿ�ȷ��Ҫ����EXECL BOM !");
//		if (messageBox.open() == SWT.OK) {
			writeToExcel(revLineListTest);
//		}
	}


	private void writeToExcel(List<RevBean> revLineList) throws InvalidFormatException,IOException, TCException  {

		FileInputStream fileInputStream = new FileInputStream(InFileName[0]);
		wb = WorkbookFactory.create(fileInputStream);
		sheet1 = wb.getSheetAt(0);
		sheet2 = wb.getSheetAt(1);
		sheet3 = wb.getSheetAt(2);
		sheet4 = wb.getSheetAt(3);

		try{
			Collections.sort(revLineList,new Comparator<RevBean>(){
				@Override
				public int compare(RevBean arg0, RevBean arg1)
				{
					int sortRes=0;
					try {
						String order_number0=arg0.OrderRevLine.getProperty("nc8_order_number");
						String order_line_number0=arg0.OrderRevLine.getProperty("nc8_order_line_number");
						String order_number1=arg1.OrderRevLine.getProperty("nc8_order_number");
						String order_line_number1=arg1.OrderRevLine.getProperty("nc8_order_line_number");
						if (order_number0==null) return -1;
						sortRes=order_number0.compareTo(order_number1);
						if (sortRes==0){
							if (order_line_number0==null) return -1;
							if (order_line_number1==null) return 1;
							String intRegex="\\d+";
							if (order_line_number0.matches(intRegex) && order_line_number1.matches(intRegex)){
								Integer line0=Integer.parseInt(order_line_number0);
								Integer line1=Integer.parseInt(order_line_number1);
								return line0.compareTo(line1);
							}
							return order_line_number0.compareTo(order_line_number1);
						}
						
					} catch (TCException e) {
						e.printStackTrace();
					}
					return sortRes;
				}
			});
		}catch(Exception e){System.out.println("�������");}
		
		/**
		 * ��ƷͼֽĿ¼
		 * �����ͺ�Ϊnll 
		 */
		// ��Ʒ�ͺ�/��Ʒͼ��(��ѡ��Ϊ����������д�����ͺţ���ѡ�в�Ϊ����������дͼ�� ͨ���������ж��Ƿ�Ϊ����)
//		String nc8_value_code = tcItemrev.getProperty("nc8_value_code");
//		boolean isWhole = nc8_value_code.startsWith("11");
//		if (isWhole) {
//			System.out.println("��ѡ�е���������");
//			String nc8_model_no = "";
//			Boolean isValid = tcItemrev.isValidPropertyName("nc8_model_no");
//			if (isValid) {
//				nc8_model_no = tcItemrev.getProperty("nc8_model_no");
//				System.out.println("nc8_model_no--------------" + nc8_model_no);
//			}else {
//				System.out.println("����������nc8_model_no");
//			}
//			//�ͺ�Ϊ��   ����ͼ��
//			if ("".equals(nc8_model_no) || nc8_model_no == null) {
//				String nc8_drawing_no = "";
//				Boolean isValid2 = tcItemrev.isValidPropertyName("nc8_drawing_no");
//				if (isValid2) {
//					nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
//					System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
//				}else {
//					System.out.println("����������nc8_drawing_no");
//				}
//				//ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_drawing_no);
//				ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_drawing_no);
//				ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_drawing_no);
//				ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_drawing_no);
//			}else {
//				//ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_model_no);
//				ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_model_no);
//				ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_model_no);
//				ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_model_no);
//			}
//		} else {
//			System.out.println("��ѡ�еĲ���������");
//			String nc8_drawing_no = "";
//			Boolean isValid = tcItemrev.isValidPropertyName("nc8_drawing_no");
//			if (isValid) {
//				nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
//				System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
//			}else {
//				System.out.println("����������nc8_drawing_no");
//			}
//			//ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_drawing_no);
//			ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_drawing_no);
//			ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_drawing_no);
//			ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_drawing_no);
//		}
//
//		// ��Ʒ����
//		String object_name_sel = tcItemrev.getProperty("object_name");
//		System.out.println("object_name--------------" + object_name_sel);
//		//ngc_utils.DoExcel.FillCell(sheet1, "H2", object_name_sel);
//		ngc_utils.DoExcel.FillCell(sheet2, "B2", object_name_sel);
//		ngc_utils.DoExcel.FillCell(sheet3, "B2", object_name_sel);
//		ngc_utils.DoExcel.FillCell(sheet4, "C9", object_name_sel);
//		
//		
//		// ��Ʒͼ��
//		String top_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
//		System.out.println("top_drawing_no--------------" + top_drawing_no);
//		ngc_utils.DoExcel.FillCell(sheet4, "C10", top_drawing_no);
		
		
		TCComponentItemRevision orderItemRevision = null;
		if(revLineList.size()>0){
			orderItemRevision = revLineList.get(0).OrderRevLine;
		}
		 
		if (orderItemRevision != null) {
			System.out.println("��������" + orderItemRevision.getProperty("object_name"));
			TCComponentItem item = orderItemRevision.getItem();
			TCComponentItemRevision latestItemRevision = item.getLatestItemRevision();

			// ���۶�����
			temp_nc8_order_number = latestItemRevision.getProperty("nc8_order_number");//������
			System.out.println("������ nc8_order_number=" + temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet2, "B3", temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet1, "M1", temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet3, "B3", temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet4, "C11", temp_nc8_order_number);

			// ����� 
			String nc8_model_no = latestItemRevision.getProperty("nc8_mo_number");
			System.out.println("nc8_mo_number--------------" + nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet2, "E3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet1, "M2", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet3, "D3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet4, "C12", nc8_model_no);

		}else{
			
			System.out.println("δ��ȡ������--------------------------------");
			
			
		}

		// ��Ʒ��ϸ��

		for (int i = 0; i < revLineList.size(); i++) {

			RevBean revbean = revLineList.get(i);
			TCComponentItemRevision revLine = revbean.RevLine;
			TCComponentItemRevision orderRevLine = revbean.OrderRevLine;
			String sequenceNo = revbean.sequenceNo;//Bom���� �㼶
			String parentMaterialCode = revbean.parentMaterialCode;//Bom���� �������ϱ���
			String bomQuantity = revbean.bomQuantity;//Bom���� ����
			String bomRemark = revbean.bomRemark;//Bom���� ��ע
			String woodPattern = revbean.woodPattern;//Bom���� ľģͼ��
//			TCComponentItemRevision bomLineRevision = bomLine.getItemRevision();
			Integer level = revbean.Level;
			String nc8_material_code_check = revLine.getProperty("nc8_material_code");
			System.out.println("��name = " + revLine.getProperty("object_name") + ", object_type = "+ revLine.getTCProperty("object_type").getStringValue() + ", " + "���ϱ��� = " + nc8_material_code_check + "��");
			if (!"".equals(nc8_material_code_check) && nc8_material_code_check != null) {
				if (nc8_material_code_check.startsWith("13") || (nc8_material_code_check.startsWith("11") && !nc8_material_code_check.substring(4, 6).equals("00"))) {
					// ���ϱ�����13��ͷΪ��� ��11��ͷ������λ��Ϊ00��Ϊ��װ
					String nc8_firstused_products = revLine.getProperty("nc8_firstused_products");
					System.out.println("��name = " + revLine.getProperty("object_name") + ", �״����ڲ�Ʒ����ֵΪ" + nc8_firstused_products + "��");
					if ("".equals(nc8_firstused_products) || nc8_firstused_products == null) {
						TCComponentUser tCComponentUserBomLine = (TCComponentUser) revLine.getRelatedComponent("owning_user");
						String owning_user = tCComponentUserBomLine.getUserId();
						System.out.println("��ǰbomLine������=======================" + owning_user);
						if (owning_user.equals(session.getUser().getUserId())) {
							MessageBox.post(revLine.getProperty("object_name")+"���״����ڲ�Ʒ��ֵΪ�գ�", "����", 1);
							return;
						} else {
							fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);

						}
					} else {
						//if (nc8_firstused_products.equals(whole_nc8_drawing_no)) {
							fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
						//} else {
							//System.out.println("���״����ڲ�Ʒ������ֵ��ͼ�Ų�ͬ");
						//}
					}
				} else {
					fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
				}
			} else {
				fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
			}
		}
		System.out.println("д������ֵ��Excel���------------- ");

		// �洢�ļ����ƣ���ǰ���ڣ�
		String time = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
		if (InFileName[0].endsWith(".xls")) {
			OutFileName = TempPath + time + ".xls";
		}
		if (InFileName[0].endsWith(".xlsx")) {
			OutFileName = TempPath + time + ".xlsx";
		}

		String saveToTCFileName = "������ϸ��";

		String temp_revision_id = "A";

		// ����Ϊ����ѡ���ͼ��+�汾��

		// ����Ϊ����EBOM��+��_��+"��ѡ�ṹ����ͼ��"+��_��+"�汾��"+"��λ��ˮ��"
//		String nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
//
//		if (nc8_drawing_no == null || nc8_drawing_no.trim().equals("")) {
//			if (!temp_nc8_order_number.equals("") || temp_nc8_order_number != null || !nc8_order_line_number.equals("") || nc8_order_line_number != null) {
//				saveToTCFileName = "EBOM" + "_" + orderNo + "_" + nc8_drawing_no + "_" + tcItemrev.getProperty("item_revision_id");
//			}else {
//				saveToTCFileName = "EBOM" + "_" + orderNo + "_" + tcItemrev.getProperty("item_revision_id");				
//			}
//
//			temp_revision_id = tcItemrev.getProperty("item_revision_id");
//			
//			
//			MessageBox.post("��ǰѡ��BOMLineͼ��Ϊ�գ�������������ʾ�쳣��", "��ʾ", MessageBox.WARNING);
//			
//
//		} else {
//
//			temp_revision_id = tcItemrev.getProperty("item_revision_id") + getSequenceCode(tcItemrev.getProperty("nc8_material_code"));
//
//			if (!temp_nc8_order_number.equals("") || temp_nc8_order_number != null || !nc8_order_line_number.equals("") || nc8_order_line_number != null) {
//				saveToTCFileName = "EBOM" + "_" + orderNo + "_" + nc8_drawing_no + "_" + temp_revision_id;
//			}else {
//				saveToTCFileName = "EBOM" + "_" + nc8_drawing_no + "_" + temp_revision_id;			
//			}
//
//		}
		// �洢�ļ����ƣ���ǰ���ڣ�
		String timeNow = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
		
		saveToTCFileName = "EBOM" + "_" + orderNo + "_" +timeNow;

		// ����汾�� �洢����ʱ��İ汾��
		System.out.println("����汾�� =" + temp_revision_id);
		ngc_utils.DoExcel.FillCell(sheet1, "O1", temp_revision_id);
		//����
		ngc_utils.DoExcel.FillCell(sheet1, "O2", userStr);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G1", temp_revision_id);
		System.out.println("���ɵ�saveToTCFileName��------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F1", temp_revision_id);
		System.out.println("���ɵ�saveToTCFileName��------------- " + saveToTCFileName);

		// ����ҳ��������������ÿҳ�ĵ�Ԫ��������
		int pageBum = 1;
		if (revLineList != null && revLineList.size() > 0) {
			pageBum = revLineList.size() / 34;
			if (pageBum <= 0) {
				pageBum = 1;
			}
		}

		System.out.println("����ҳ�� =" + temp_revision_id);
		// sheet1
//		ngc_utils.DoExcel.FillCell(sheet1, "O2", pageBum + "");
//		System.out.println("����ҳ����------------- " + saveToTCFileName);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G2", pageBum + "");
		System.out.println("sheet2����ҳ����------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F2", pageBum + "");
		System.out.println("sheet3����ҳ����------------- " + saveToTCFileName);

		// �洢�ļ�����
		if (InFileName[0].endsWith(".xls")) {
			OutFileName = TempPath + saveToTCFileName + ".xls";
		}
		if (InFileName[0].endsWith(".xlsx")) {
			OutFileName = TempPath + saveToTCFileName + ".xlsx";
		}

		// д���ļ�������
		FileOutputStream fileOut = new FileOutputStream(OutFileName);
		wb.write(fileOut);
		fileOut.close();

		System.out.println("OutFileName--------------" + OutFileName);
		
		String queryLineNumber = "10";
		while (true) {
			TCComponent[] topComponent = Common.CommonFinder("007-����", "nc8_order_number,nc8_order_line_number", orderNo+","+queryLineNumber);
//			TCComponent[] topComponent = Common.CommonFinder("007-����", "nc8_order_number,nc8_order_line_number", orderNo+","+"10");
			if (null != topComponent){
				if(topComponent[0] instanceof TCComponentItemRevision){
					TCComponentItemRevision orderItemRev = (TCComponentItemRevision) topComponent[0];
					String item_revision_id = orderItemRev.getProperty("item_revision_id");
					String item_id = orderItemRev.getProperty("item_id");
					String nc8_material_code = orderItemRev.getProperty("nc8_material_code");
					if("".equals(nc8_material_code)){
						MessageBox.post("����"+orderNo+"-"+queryLineNumber+"û�����ϱ��룡","����",MessageBox.WARNING);
					}
					
					// TODO
					TCComponentItemRevision tccir = getItemRevision("����ĵ��汾", "EBOM" + "_" + orderNo +"��ϸ��", orderItemRev, "NC8_other_deliverables_rel");
					if (null != tccir) {
						String NC8_design_doc_revision = tccir.getProperty("item_revision_id");
						if(!item_revision_id.equals(NC8_design_doc_revision)){
							tccir=tccir.saveAs(item_revision_id); //�������
							orderItemRev.add("NC8_other_deliverables_rel", tccir);
						}
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
					}else {
						TCComponentItem item = null;
						if("A".equals(item_revision_id)){
							TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
							item = itemType.create(itemType.getNewID(),item_revision_id, "NC8_design_doc", "EBOM" + "_" + orderNo +"��ϸ��" , "", null);
							TCComponentItemRevision itemRevision = item.getLatestItemRevision();
							itemRevision.setProperty("nc8_business_unit", "IBD");
							itemRevision.setProperty("nc8_small_class", "EBOM");
							itemRevision.setProperty("nc8_subclass", "EBOM");
							if (itemRevision.isValidPropertyName("nc8_material_code")) {	//���Ϸ�������Ϊд�����ʱ��û�����������
								itemRevision.setProperty("nc8_material_code", nc8_material_code);
							} 
							String nc8_document_num2 = generateNumber(itemRevision);
							if("".equals(nc8_document_num2)){
								MessageBox.post("���ɵ��ĵ���Ų��ɹ���","����",MessageBox.ERROR);
							}else{
								itemRevision.setProperty("nc8_document_num2", nc8_document_num2);
							}
							orderItemRev.add("NC8_other_deliverables_rel", itemRevision);
							TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
//						itemRevision.add("IMAN_reference", datasetComponent);
							itemRevision.add("IMAN_specification", datasetComponent);
						}else{
							char new_item_revision_id = item_revision_id.toCharArray()[0];
							new_item_revision_id = (char) (new_item_revision_id - 1) ;
							String query_revision_id = null;
							TCComponent[] componentzj = null;
							TCComponentItemRevision need_update_itemrevision = null;
							//��������ĵ��汾�汾����
							do{
								query_revision_id = Character.toString(new_item_revision_id);
								componentzj = Common.CommonFinder("������汾...", "ItemID,Revision", item_id + "," + query_revision_id);
								if(null != componentzj){
									need_update_itemrevision = getItemRevision("����ĵ��汾", "EBOM" + "_" + orderNo +"��ϸ��", ((TCComponentItemRevision)componentzj[0]), "NC8_other_deliverables_rel");
									if(null != need_update_itemrevision){
										break;
									}
								}
								new_item_revision_id = (char) (new_item_revision_id - 1) ;
								
							}while("A".compareTo(Character.toString(new_item_revision_id))<=0);
							
							if(null == need_update_itemrevision){//ֱ�Ӵ���һ���붩���汾һ�µ�itemrevision
								TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
								item = itemType.create(itemType.getNewID(),item_revision_id, "NC8_design_doc", "EBOM" + "_" + orderNo +"��ϸ��" , "", null);
								TCComponentItemRevision itemRevision = item.getLatestItemRevision();
								itemRevision.setProperty("nc8_business_unit", "IBD");
								itemRevision.setProperty("nc8_small_class", "EBOM");
								itemRevision.setProperty("nc8_subclass", "EBOM");
								if (itemRevision.isValidPropertyName("nc8_material_code")) {	//���Ϸ�������Ϊд�����ʱ��û�����������
									itemRevision.setProperty("nc8_material_code", nc8_material_code);
								} 
								String nc8_document_num2 = generateNumber(itemRevision);
								if("".equals(nc8_document_num2)){
									MessageBox.post("���ɵ��ĵ���Ų��ɹ���","����",MessageBox.ERROR);
								}else{
									itemRevision.setProperty("nc8_document_num2", nc8_document_num2);
								}
								orderItemRev.add("NC8_other_deliverables_rel", itemRevision);
								TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
//							itemRevision.add("IMAN_reference", datasetComponent);
								itemRevision.add("IMAN_specification", datasetComponent);
							}else{//��������
								TCComponentItemRevision update_itemrevision = need_update_itemrevision.saveAs(item_revision_id); //�������
								orderItemRev.add("NC8_other_deliverables_rel", update_itemrevision);
								TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
								update_itemrevision.add("IMAN_specification", datasetComponent);
							}
							
						}
					}
					
					// ���ļ�д��TC���ڸö����ţ������к�Ϊ10������������α�ļ�������
//				createOrUpdateExcel(OutFileName, saveToTCFileName, orderItemRev, "NC8_other_deliverables_rel", true);
					System.out.println("д��TC�ɹ�--------------");
					MessageBox.post("���ɱ�����ϸ��ɹ��������"+orderNo+"-"+queryLineNumber+"��������","��Ϣ",MessageBox.INFORMATION);
					break;
				}
			}else {
				queryLineNumber = (Integer.parseInt(queryLineNumber) + 10) + "";
			}
			
			if (Integer.parseInt(queryLineNumber) > 1000000) {
				break;
			}
			
		}
//		else{
//			TCComponent[] topComp = Common.CommonFinder("007-����", "nc8_order_number", orderNo);
//			if(topComp!=null){
//				if(topComp[0] instanceof TCComponentItemRevision){
//					TCComponentItemRevision orderItemRev = (TCComponentItemRevision) topComp[0];
//					String nc8_order_line_number = orderItemRev.getProperty("nc8_order_line_number");
//					// ���ļ�д��TC���ڸö����ţ������к�Ϊnc8_order_line_number������������α�ļ�������
//					createOrUpdateExcel(OutFileName, saveToTCFileName, orderItemRev, "NC8_other_deliverables_rel", true);
//					System.out.println("д��TC�ɹ�--------------");
//				}
//			}
//		}

		

		// ��Ԥ��
		//Runtime.getRuntime().exec("cmd /c start " + OutFileName);
//		MessageBox.post("����������ϣ�", "��ʾ", 2);

	}

	private void fillValue(TCComponentItemRevision revLine, TCComponentItemRevision orderRevLine,Sheet sheet12,Integer level, Sheet sheet22,String sequenceNo,String parentMaterialCode,String bomQuantity,String bomRemark,String woodPattern) throws TCException {
		System.out.println("row = " + rowNum + ", number = " + number + ", level = " + level + ", productRowNum = " + productRowNum);
		// ����
		String daihao = "";
		// �㼶
		String bl_sequence_no = "";
		// �汾
		String item_revision_id = "";
		// ��������
		String object_name = "";
		// Ӣ������
		String nc8_part_name = "";
		// ���ϱ���
		String nc8_material_code = "";
		// �������
		//String nc8_material_code_parent = "";
		// ����
		String nc8_material = "";
		// ����
		String bl_quantity = "";
		// ����
		String nc8_weight = "";
		// ��ע
		String NC8_BOM_remark = "";
		// ľģͼ��
		String nc8_wood_pattern = "";
		// �����к�
		String nc8_order_line_number = "";

		String object_type = revLine.getTCProperty("object_type").getStringValue();
		System.out.println("��object_type��" + object_type);
		if ("NC8_cust_supplyRevision".equals(object_type) || "NC8_CastingRevision".equals(object_type) || "NC8_ForgingsRevision".equals(object_type)
				|| "NC8_WeldingRevision".equals(object_type) || "NC8_SectionRevision".equals(object_type)
				|| "NC8_AssistantMatRevision".equals(object_type) || "NC8_test_piecesRevision".equals(object_type)
				|| "NC8_purchasedRevision".equals(object_type)) {

			if ("NC8_cust_supplyRevision".equals(object_type)) {
				/**
				 * �͹���
				 */
				System.out.println("���ö���Ϊ�͹�����");
				// ���� nc8_drawing_no+�� ��+nc8_specification
				String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
				System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
				String nc8_specification = revLine.getProperty("nc8_specification");
				System.out.println("�����--------------" + nc8_specification);
				String drawing_no3 = revLine.getProperty("nc8_drawing_no3");
				//daihao = nc8_drawing_no + " " + nc8_specification;
				daihao = drawing_no3;//2019/01/15���ģ���nc8_Standard��nc8_Specificationƴ�ӵ�ֵ��Ϊnc8_drawing_no3��ֵ
				System.out.println("�����š�--------------" + daihao);
				// �㼶
				//bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				//System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = revLine.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// �����к�
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("�������кš�--------------" + nc8_order_line_number);
				// Ӣ������
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = revLine.getProperty("nc8_material_code");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					}
//
//				} else {
//					System.out.println("��û���ϲ��������");
//					nc8_material_code_parent = "";
//				}

				// ���� (����δд��)

				/**
				 * nc8_order_number =
				 * bomLineRevision.getProperty("nc8_order_number");
				 * System.out.println("�����ϡ�--------------" + nc8_drawing_no);
				 */

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
//				bl_quantity = bomLine.getProperty("bl_quantity");
//				System.out.println("��������--------------" + bl_quantity);
//				if (bl_quantity == null || bl_quantity.length() == 0) {
//					bl_quantity = "1";
//				}

				// ����
				nc8_weight = revLine.getProperty("nc8_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);

				// ľģͼ�� item ���ԣ��о��ã�
//				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
//					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
//					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
//				}

				// ��������

			} else if ("NC8_CastingRevision".equals(object_type) || "NC8_ForgingsRevision".equals(object_type)
					|| "NC8_WeldingRevision".equals(object_type) || "NC8_SectionRevision".equals(object_type)) {
				/**
				 * ԭ����
				 */
				System.out.println("���ö���Ϊԭ���ϡ�");
				// ���� nc8_Standard+�� ��+nc8_specification
				String nc8_Standard = revLine.getProperty("nc8_Standard");
				System.out.println("����׼��--------------" + nc8_Standard);
				String nc8_Specification = revLine.getProperty("nc8_Specification");
				System.out.println("�����--------------" + nc8_Specification);
				String drawing_no3 = revLine.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_Specification;
				daihao = drawing_no3;//2019/01/15���ģ���nc8_Standard��nc8_Specificationƴ�ӵ�ֵ��Ϊnc8_drawing_no3��ֵ
				System.out.println("�����š�--------------" + daihao);
				// �㼶
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = revLine.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// �����к�
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("�������кš�--------------" + nc8_order_line_number);
				// Ӣ������
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = revLine.getProperty("nc8_Materialnumber");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("��û���ϲ��������");
//					nc8_material_code_parent = "";
//				}

				// ����
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
//				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("��������--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// ����
				nc8_weight = revLine.getProperty("nc8_net_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			} else if ("NC8_AssistantMatRevision".equals(object_type)) {
				/**
				 * ����
				 */
				System.out.println("���ö���Ϊ���ϡ�");
				// ���� nc8_Standard +�� ��+ nc8_model+�� ��+ nc8_Specification
				String nc8_Standard = revLine.getProperty("nc8_Standard");
				System.out.println("����׼��--------------" + nc8_Standard);
				String nc8_model = revLine.getProperty("nc8_model");
				System.out.println("���ͺš�--------------" + nc8_model);
				String nc8_specification = revLine.getProperty("nc8_Specification");
				System.out.println("�����--------------" + nc8_specification);
				String drawing_no3 = revLine.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_model + " " + nc8_specification;
				daihao = drawing_no3;//2019/01/15���ģ���nc8_Standard��nc8_Specificationƴ�ӵ�ֵ��Ϊnc8_drawing_no3��ֵ
				System.out.println("�����š�--------------" + daihao);
				// �㼶
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = revLine.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// �����к�
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("�������кš�--------------" + nc8_order_line_number);
				// Ӣ������
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = revLine.getProperty("nc8_Materialnumber");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("��û���ϲ��������");
//					nc8_material_code_parent = "";
//				}

				// ����
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
				/**
				 * ���Ȼ�ȡ������������������Ϊ�յ�ʱ����ȥ��ȡ����
				 */
//				String nc8_assist_number = bomLine.getProperty("NC8_Assist_number");
//				System.out.println("���������� = " + nc8_assist_number + "��");
//				if ("".equals(nc8_assist_number) || nc8_assist_number == null) {
//					String bl_quantity_bak = bomLine.getProperty("bl_quantity");
//					System.out.println("��������--------------" + bl_quantity_bak);
//					if (bl_quantity_bak == null || bl_quantity_bak.length() == 0) {
//						bl_quantity = "1";
//					}else {
//						bl_quantity = bl_quantity_bak;
//					}
//				}else {
//					bl_quantity = nc8_assist_number;
//				}
				System.out.println("��excel������--------------" + bl_quantity);
				

				// ����
				nc8_weight = revLine.getProperty("nc8_net_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			} else if ("NC8_test_piecesRevision".equals(object_type)) {
				/**
				 * �����
				 */
				System.out.println("���ö���Ϊ�������");
				// ���� nc8_drawing_no
				String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
				System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("�����š�--------------" + daihao);
				// �㼶
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = revLine.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// �����к�
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("�������кš�--------------" + nc8_order_line_number);
				// Ӣ������
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = revLine.getProperty("nc8_material_code");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("��û���ϲ��������");
//					nc8_material_code_parent = "";
//				}

				// ����
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
//				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("��������--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// ����
				nc8_weight = revLine.getProperty("nc8_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			} else if ("NC8_purchasedRevision".equals(object_type)) {
				/**
				 * �⹺��
				 */
				System.out.println("���ö���Ϊ�⹺����");
				// ���� nc8_drawing_no
				String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
				System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("�����š�--------------" + daihao);
				// �㼶
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = revLine.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// �����к�
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("�������кš�--------------" + nc8_order_line_number);
				// Ӣ������
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = revLine.getProperty("nc8_material_code");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("��������롿--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("��û���ϲ��������");
//					nc8_material_code_parent = "";
//				}

				/**
				 * ���� nc8_material+nc8_grade+nc8_hardness_level 2018-08-21�޸�
				 * 1.��ϸ�����⹺��ITEM�ġ����ϡ���ȡITEM���ԣ�����+���ܵȼ�+Ӳ�ȵȼ���
				 * 2.��ϸ�����⹺��ITEM�ġ���ע����ȡITEM���ԣ�������+BOM��ע��
				 */
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("�����ʡ�--------------" + nc8_material);
				String nc8_grade = revLine.getProperty("nc8_grade");
				System.out.println("�����ܵȼ���--------------" + nc8_grade);
				String nc8_hardness_level = revLine.getProperty("nc8_hardness_level");
				System.out.println("��Ӳ�ȵȼ���--------------" + nc8_hardness_level);
				nc8_material = nc8_material + " " + nc8_grade + " " + nc8_hardness_level;
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
//				bl_quantity = bomLine.getProperty("bl_quantity");
//				System.out.println("��������--------------" + bl_quantity);
//				if (bl_quantity == null || bl_quantity.length() == 0) {
//					bl_quantity = "1";
//				}
				// ����
				nc8_weight = revLine.getProperty("nc8_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע nc8_feature_set- nc8_grade - nc8_hardness_level
				// +NC8_BOM_remark
				String nc8_feature_set = revLine.getProperty("nc8_feature_set");
				System.out.println("����������--------------" + nc8_feature_set);

				/**
				 * String nc8_grade = bomLineRevision.getProperty("nc8_grade");
				 * System.out.println("�����ܵȼ���--------------" + nc8_grade);
				 * String nc8_hardness_level =
				 * bomLineRevision.getProperty("nc8_hardness_level");
				 * System.out.println("��Ӳ�ȵȼ���--------------" +
				 * nc8_hardness_level);
				 */

				/**
				 * 08/20����� EBOM���⹺���ı�ע����ԭ�������: BOm��ע ����Ϊ��������+
				 * BOM��ע���Կո����ӣ�����ֵ���ݣ����������ӷ���
				 */
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("��BOM��ע��--------------" + NC8_BOM_remark);
				NC8_BOM_remark = (nc8_feature_set + " " + NC8_BOM_remark).trim();
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			}

		} else {
			System.out.println("���ö���Ϊ��ͨ����");
			// ���� nc8_drawing_no
			String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
			System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
			daihao = nc8_drawing_no;
			System.out.println("�����š�--------------" + daihao);
			// �㼶
//			bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//			System.out.println("���㼶��" + bl_sequence_no);
			// �汾
			item_revision_id = revLine.getProperty("item_revision_id");
			System.out.println("���汾��--------------" + item_revision_id);
			// ��������
			object_name = revLine.getProperty("object_name");
			System.out.println("���������ơ�--------------" + object_name);
			// �����к�
			nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
			System.out.println("�������кš�--------------" + nc8_order_line_number);
			// Ӣ������
			nc8_part_name = revLine.getProperty("nc8_part_name");
			System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
			// ���ϱ���
			nc8_material_code = revLine.getProperty("nc8_material_code");
			System.out.println("�����ϱ��롿--------------" + nc8_material_code);
			// �������
//			TCComponentBOMLine parentBomLine = bomLine.parent();
//			if (parentBomLine != null) {
//				TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//				if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//					System.out.println("��������롿--------------" + nc8_material_code_parent);
//				} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//					System.out.println("��������롿--------------" + nc8_material_code_parent);
//				}
//			} else {
//				System.out.println("��û���ϲ��������");
//				nc8_material_code_parent = "";
//			}

			// ���� nc8_material+nc8_grade+nc8_hardness_level
			nc8_material = revLine.getProperty("nc8_material");
			System.out.println("�����ϡ�--------------" + nc8_material);
			// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
//			bl_quantity = bomLine.getProperty("bl_quantity");
//			System.out.println("��������--------------" + bl_quantity);
//			if (bl_quantity == null || bl_quantity.length() == 0) {
//				bl_quantity = "1";
//			}
			// ����
			nc8_weight = revLine.getProperty("nc8_weight");
			System.out.println("�����ء�--------------" + nc8_weight);
			// ��ע
//			NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
			System.out.println("����ע��--------------" + NC8_BOM_remark);
			// ľģͼ��
			if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
				nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
				System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
			}

			// ��������

		}
		// ��ţ���ˮ��������
		ngc_utils.DoExcel.FillCell(sheet1, "A" + rowNum, number + "");
		String blankStr = "";

		if(level == 10000){//��������levelΪ10000������û��Bom�ṹ�����Բ����ڲ㼶���㼶Ϊ��
			ngc_utils.DoExcel.FillCell(sheet1, "B" + rowNum, "");
		}else{
			if (level > 1) {
				for (int j = 0; j < level - 1; j++) {
					blankStr = blankStr + "  ";
				}
			}
			bl_sequence_no = blankStr + "L" + (level - 1);
		}
		
		ngc_utils.DoExcel.FillCell(sheet1, "B" + rowNum, bl_sequence_no);//BomLine��ȡ ��������//bl_sequence_no
		ngc_utils.DoExcel.FillCell(sheet1, "C" + rowNum, daihao);
		ngc_utils.DoExcel.FillCell(sheet1, "D" + rowNum, item_revision_id);
		ngc_utils.DoExcel.FillCell(sheet1, "E" + rowNum, object_name);
		ngc_utils.DoExcel.FillCell(sheet1, "F" + rowNum, nc8_part_name);
		ngc_utils.DoExcel.FillCell(sheet1, "G" + rowNum, nc8_material_code);
		ngc_utils.DoExcel.FillCell(sheet1, "H" + rowNum, parentMaterialCode);//BomLine�ĸ�
		ngc_utils.DoExcel.FillCell(sheet1, "I" + rowNum, nc8_material);
		ngc_utils.DoExcel.FillCell(sheet1, "J" + rowNum, bomQuantity);//BomLine��ȡ
		ngc_utils.DoExcel.FillCell(sheet1, "K" + rowNum, nc8_weight);
		ngc_utils.DoExcel.FillCell(sheet1, "L" + rowNum, bomRemark);//BomLine��ȡ
		ngc_utils.DoExcel.FillCell(sheet1, "M" + rowNum, woodPattern);//BomLine��ȡ
		ngc_utils.DoExcel.FillCell(sheet1, "N" + rowNum, "");
		// System.out.println("�����ղ㼶��" + bl_sequence_no);
		ngc_utils.DoExcel.FillCell(sheet1, "P" + rowNum, nc8_order_line_number);
		rowNum++;
		number++;

		TCComponent[] relatedComponents = revLine.getRelatedComponents("IMAN_specification");
		System.out.println("��������item��size��" + relatedComponents.length);
		Boolean isCreate = false;
		for (int j = 0; j < relatedComponents.length; j++) {
			TCComponent tcComponent = relatedComponents[j];
			String string = tcComponent.getProperty("object_name");
			System.out.println("��object_name��" + string);
			if (tcComponent instanceof TCComponentDataset) {
				TCComponentDataset dataset = (TCComponentDataset) tcComponent;
				String objectType = dataset.getProperty("object_type");
				System.out.println("���������� object_type = " + objectType + "��");
				// UGPART , UGMASTER, PDF, TIF
				if (objectType.equals("UGPART") || objectType.equals("PDF") || objectType.equals("TIF")) {
					isCreate = true;
					break;
				}
			}

		}
		
//		boolean isRoot = bomLine.isRoot();
		
//		if (isCreate || isRoot) {
		if (isCreate) {
			/**
			 * ��ƷͼֽĿ¼
			 */
			// ͼ��
			String nc8_drawing_no_product = revLine.getProperty("nc8_drawing_no");
			System.out.println("����ƷͼֽĿ¼-ͼ�š�--------------" + nc8_drawing_no_product);
			ngc_utils.DoExcel.FillCell(sheet2, "A" + productRowNum, nc8_drawing_no_product);

			// �汾
			String item_revision_id_product = revLine.getProperty("item_revision_id");
			System.out.println("����ƷͼֽĿ¼-�汾��--------------" + item_revision_id_product);
			ngc_utils.DoExcel.FillCell(sheet2, "B" + productRowNum, item_revision_id_product);

			// ��������
			String object_name_product = revLine.getProperty("object_name");
			System.out.println("����ƷͼֽĿ¼-�������ơ�--------------" + object_name_product);
			ngc_utils.DoExcel.FillCell(sheet2, "C" + productRowNum, object_name_product);

			// Ӣ������
			String nc8_part_name_product = revLine.getProperty("nc8_part_name");
			System.out.println("����ƷͼֽĿ¼-Ӣ�����ơ�--------------" + nc8_part_name_product);
			ngc_utils.DoExcel.FillCell(sheet2, "D" + productRowNum, nc8_part_name_product);

			// ͼ��
			String nc8_drawing_size_product = revLine.getProperty("nc8_drawing_size");
			System.out.println("����ƷͼֽĿ¼-ͼ����--------------" + nc8_drawing_size_product);
			ngc_utils.DoExcel.FillCell(sheet2, "E" + productRowNum, nc8_drawing_size_product);

			// ҳ��
			String nc8_pages_product = revLine.getProperty("nc8_pages");
			System.out.println("����ƷͼֽĿ¼-ҳ����--------------" + nc8_pages_product);
			ngc_utils.DoExcel.FillCell(sheet2, "F" + productRowNum, nc8_pages_product);

			// ��ע
			String nc8_remarks_product = revLine.getProperty("nc8_remarks");
			System.out.println("����ƷͼֽĿ¼-��ע��--------------" + nc8_remarks_product);
			ngc_utils.DoExcel.FillCell(sheet2, "G" + productRowNum, nc8_remarks_product);

			productRowNum++;
		}			
	}


	private String getSequenceCode(String sequenceName) {

		if (sequenceName != null && sequenceName.length() != 0) {

			System.out.println("�ϴ���sequenceNameΪ------------- " + sequenceName);
			String sequenceCode = String.valueOf(JDBCUtils.querySequenceCode(sequenceName));
			System.out.println("������ˮ��ֵΪ------------- " + sequenceCode);

			if (sequenceCode.equals("-1")) {
				MessageBox.post("��ˮ���ȡʧ��!", "��ʾ", MessageBox.WARNING);
				throw new RuntimeException("��ȡ��ˮ��ʧ��");
			} else {
				while (sequenceCode.length() < 2) {
					StringBuilder sb = new StringBuilder("0");
					sb.append(sequenceCode);
					sequenceCode = sb.toString();
				}
				return sequenceCode;
			}

		} else {

			MessageBox.post("ѡ��BOM�����ϱ��벻��Ϊ�գ�", "����", 1);
			return "";
		}

	}
	
	
	
	
	//2018/12/28��
	
	//��ngc_utils.ReportCommon�п�����������ΪsetDatasetFileToTC������MessageDialog.openQuestion�����ᵯ�����Ѵ��ڣ��Ƿ񸲸�ʱ�����ڱ�����û��Ч��
	//����Ϊnull argument:The dialog should be created in UI thread ����Ч��ԭ���������Ϊ�������Ѿ�����һ��swing���档
	
	//����MessageDialog.openQuestion�ĳ�JOptionPane.showOptionDialog�Ϳ����ˣ���������������������õ�����ʱ���ͻῨ��������ΪSWT UI Thread is not responding!
	
	//�� ���ϴ�Excel���ݼ���TC�ķ���������һ�ݵ�����������û�뵽���õİ취����ʱ��ôд
//======================================================================================================================================================================	
	/**
	 * �ϴ�Excel���ݼ���TC
	 * 
	 * @param localFile �����ļ���
	 * @param datasetName �ϴ������ݼ�����
	 * @param itemRevision �ҿ�����
	 * @param relationType �ҿ�����
	 * @param replaceAlert �Ƿ��滻
	 * @return
	 */
	public synchronized static TCComponentDataset createOrUpdateExcel(String localFile, String datasetName, 
			TCComponent relationObject, String relationType, boolean replaceAlert) {
		try {
			String datasetType = "";
			if(localFile.endsWith(".xls")){
				datasetType = "MS Excel";
			}
			if(localFile.endsWith(".xlsx")  || localFile.endsWith(".xlsm")){
				datasetType = "MS ExcelX";
			}
			TCComponentDataset datasetComponent = hasDataset(datasetType, datasetName, relationObject, relationType);
			
			if (datasetComponent == null) {
				datasetComponent = setDatasetFileToTC(localFile, datasetType, "excel", datasetName);
				relationObject.add(relationType, datasetComponent);
			} else {
				setDatasetFileToTC(localFile, datasetComponent,  "excel", replaceAlert);
			}
			return datasetComponent;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}
	
	//�ж϶������Ƿ�����ض���ϵ�����ݼ�
		public static TCComponentDataset hasDataset(String datasetType,String datasetName, TCComponent relationObject, String relationType) {
			try {
				TCComponent TCComponent[] = relationObject.getRelatedComponents(relationType);
				if ((TCComponent != null) && (TCComponent.length > 0)) {
					for (int i = 0; i < TCComponent.length; i++) {
						if (	   (TCComponent[i].getProperty("object_type").equals(datasetType))
								&& (TCComponent[i].getProperty("object_name").equals(datasetName))) {
							return (TCComponentDataset) TCComponent[i];
						}
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
				return null;
			}
			return null;
		}
		
		/**
		 * �ϴ������ļ���TC 
		 * TCϵͳ�в��������ļ������ݼ�
		 * @param localFile
		 * @param datasetType
		 * @param datasetNamedRef
		 * @param datasetName
		 * @return
		 */
		public static TCComponentDataset setDatasetFileToTC(String localFile, String datasetType, String datasetNamedRef, String datasetName) {
			try{
				TCSession tcSession = (TCSession) AIFUtility.getDefaultSession();
				String filePathNames[] = { localFile };
				String namedRefs[] = { datasetNamedRef };
				TCTypeService typeService = tcSession.getTypeService();
				TCComponentDatasetType TCDatasetType = (TCComponentDatasetType) typeService.getTypeComponent(datasetType.replace(" ", ""));
				TCComponentDataset datasetComponent = TCDatasetType.setFiles(datasetName,"Created by program.", datasetType.replace(" ", ""), filePathNames,namedRefs);
				return datasetComponent;
				
			} catch (Exception e) {
				e.printStackTrace();
				return null;
			}
		}
		
		/**
		 * �ϴ������ļ���TC 
		 * TCϵͳ���Ѻ��б����ļ������ݼ�
		 * @param localFile
		 * @param datasetComponent
		 * @param datasetNamedRef
		 * @param replaceAlert
		 *        true ��ʾ�Ƿ񸲸��ж϶Ի���
		 *        false ֱ�Ӹ��� ����ʾ
		 *        2018/12/1 ��Ϊ���Ѵ����Ƿ񸲸ǡ���ʾ��UI�̴߳��󣬹��޸���ʾ��
		 *        2018/12/28�Ѹ÷�������������������������Ϊ��ĵط��õ��˷���ʱ��JOptionPane.showOptionDialog�ᵼ�¿�������������ֻ����JOptionPane.showOptionDialog��������MessageDialog.openQuestion��
		 *        ���������ڱ������ڵ������Ѵ��ڣ��Ƿ񸲸ǡ�֮ǰ�Ѿ�����һ��Swing���棬��֪���Ƿ������ԭ��
		 */
		public static void setDatasetFileToTC(String localFile, TCComponentDataset datasetComponent, String datasetNamedRef, boolean replaceAlert) {
			try{
				TCSession tcSession = (TCSession) AIFUtility.getDefaultSession();
				String filePathNames[] = { localFile };
				String namedRefs[] = { datasetNamedRef };
				String datasetType = datasetComponent.getType();
				TCTypeService typeService = tcSession.getTypeService();
				TCComponentDatasetType TCDatasetType = (TCComponentDatasetType) typeService.getTypeComponent(datasetType);
				
				if (replaceAlert) {
//					boolean confirm = MessageDialog.openQuestion(null, "ȷ��",datasetComponent.toString() + "�Ѿ����ڣ��Ƿ񸲸ǣ�");
//					if (confirm) {
//						if (datasetComponent.getProperty("date_released").length() > 0) {
//							Common.ShowTcErrAndMsg("�����ѷ������޷��滻��");
//							return;
//						} else {
//							TCDatasetType.setFiles(datasetComponent,filePathNames, namedRefs);
//						}
//					}
					Object[] options = {"ȷ��","ȡ��"};
					int response=JOptionPane.showOptionDialog(null, datasetComponent.toString() + "�Ѿ����ڣ��Ƿ񸲸ǣ�", "ѡ��Ի������",JOptionPane.YES_OPTION,  JOptionPane.QUESTION_MESSAGE, null, options, options[0]);
					if(response==0){
						if (datasetComponent.getProperty("date_released").length() > 0) {
							Common.ShowTcErrAndMsg("�����ѷ������޷��滻��");
							return;
						} else {
							TCDatasetType.setFiles(datasetComponent,filePathNames, namedRefs);
						}
					}
				}else {
					if (datasetComponent.getProperty("date_released").length() > 0) {
						Common.ShowTcErrAndMsg("�����ѷ������޷��滻��");
						return;
					} else {
						TCDatasetType.setFiles(datasetComponent,filePathNames, namedRefs);
					}
				}
				
			} catch (Exception e) {
				e.printStackTrace();
				return;
			}
		}
	//======================================================================================================================================================================
		
		
		
		//�ж϶������Ƿ�����ض���ϵ��������汾
		public TCComponentItemRevision getItemRevision(String itemType, String itemName, TCComponent relationObject, String relationType) {
			TCComponentItemRevision tccir = null;
			try {
				TCComponent TCComponent[] = relationObject.getRelatedComponents(relationType);
				if ((TCComponent != null) && (TCComponent.length > 0)) {
					String revision = "A";
					for (int i = 0; i < TCComponent.length; i++) {
						if (	   (TCComponent[i].getProperty("object_type").equals(itemType))
								&& (TCComponent[i].getProperty("object_name").equals(itemName))) {
							String revision1 = TCComponent[i].getProperty("item_revision_id");
							if(revision.compareTo(revision1) == 0 || revision.compareTo(revision1) == (-1)){
								revision = revision1;
								tccir = (TCComponentItemRevision) TCComponent[i];
							}
						}
					}
				}
				return tccir;
			} catch (Exception e) {
				e.printStackTrace();
			}
			return null;
		}
		
		
		//��/NGC_PLM/src/com/uds/drawingNumber/codeGeneration/handlers/DocNumberGeneration.java���Ƶ�
		//------------------------------------------------------------------------------------------
			
			/**
			 * �����ĵ����
			 * �ĵ���Ź���ҵ��Ԫ-����-С��-���+4λ��ˮ
			 * @throws TCException
			 */
			private String generateNumber(TCComponentItemRevision tcItemrev) throws TCException{
				
				TCProperty type = tcItemrev.getTCProperty("object_type");
				
				String bigClass = "";
				//ҵ��Ԫ
				String bussinessUnit=tcItemrev.getProperty("nc8_business_unit");
				System.out.println("bussinessUnitΪ--------------" + bussinessUnit);
				if(bussinessUnit == null || bussinessUnit.equals("")){
					MessageBox.post("ҵ��Ԫ����Ϊ��!","��ʾ",MessageBox.WARNING);
					return "";
				}
				//��ȡҵ��Ԫǰ��λ
				String subBussinessUnit = bussinessUnit.substring(0, 3);
				System.out.println("subBussinessUnitΪ--------------" + subBussinessUnit);
				
				//���� �ж����
				if("NC8_design_docRevision".equals(type.getStringValue())){
					 bigClass= "DD";
				}else if("NC8_process_docRevision".equals(type.getStringValue())){
					 bigClass= "PD";
				}else if("NC8_standard_docRevision".equals(type.getStringValue())){
					 bigClass= "SD";
				}else if("NC8_general_docRevision".equals(type.getStringValue())){
					 bigClass= "GD";
				}
				
				
//				String bigClass=tcItemrev.getProperty("nc8_big_class");
				System.out.println("bigClassΪ--------------" + bigClass);
			   	
				
				//С��
				TCProperty smallClass = tcItemrev.getTCProperty("nc8_small_class");
//				String smallClass="VB";
				System.out.println("smallClassΪ--------------" + smallClass.getStringValue());
				
				if(smallClass.getStringValue()==null ||smallClass.getStringValue().length() == 0  ){
					
					MessageBox.post("С�಻��Ϊ�գ�","��ʾ",MessageBox.INFORMATION);
					return "";
				}
				
				
				StringBuffer sb1=new StringBuffer();
				StringBuffer sb=new StringBuffer();
				//���
				String yy =String.valueOf(Calendar.getInstance().get(Calendar.YEAR));
				sb1.append(subBussinessUnit.trim()).append(bigClass.trim()).append(smallClass.getStringValue().trim()).append(yy.substring(yy.length()-2));	
				
				
				String  tempCode = 	sb.append(subBussinessUnit.trim()).append("-").append(bigClass.trim()).append("-").append(smallClass.getStringValue().trim()).append("-").append(yy.substring(yy.length()-2)).toString();
				
				String sequenceCode =  getSequenceCodeII(sb1.toString()).toString().trim();
				
				String 	docuNum = tempCode+sequenceCode;
				
				System.out.println("�������ɵ��ĵ����Ϊ--------------" + docuNum);
				
//				try {
					//tcItemrev.setProperty("nc8_document_num2", docuNum);//����ʵ��Ϊnc8_document_num��ֻ��״̬��������nc8_document_num2Ϊ��֮��������ȵ����ԣ�BMIDE������Ϊ���ؿ�д
					//MessageBox.post("���ɵ��ĵ����Ϊ��"+docuNum,"��ʾ",MessageBox.INFORMATION);
					return docuNum;
//				} catch (Exception e1) {
//					MessageBox.post(e1.toString(),"��ʾ",MessageBox.WARNING);
//					e1.printStackTrace();
//					return "";
//				}
			}

			private String getSequenceCodeII(String sequenceName) {
				System.out.println("�ϴ���sequenceNameΪ------------- "+sequenceName);			
				String sequenceCode = String.valueOf(JDBCUtils.querySequenceCode(sequenceName));	
				System.out.println("������ˮ��ֵΪ------------- "+sequenceCode);	
				
				if(sequenceCode.equals("-1")){
					MessageBox.post("��ˮ���ȡʧ��!","��ʾ",MessageBox.WARNING);
					throw new RuntimeException("��ȡ��ˮ��ʧ��");
				}else {
			        while (sequenceCode.length()<4){
			        	StringBuilder sb=new StringBuilder("0");
			        	sb.append(sequenceCode);
			        	sequenceCode=sb.toString();
			        }
					return sequenceCode;	
				}
			}
		//------------------------------------------------------------------------------------------

		
	//	���û�ѡ�񵼳���Щ����		
	public void querySparePartsUI(ArrayList<SparePartsInfoBean> aL){
		
		final ArrayList<String> result = new ArrayList<String>();
		
		
		final JFrame jf = new JFrame("��ѡ��Ҫ�����ı���"); // ��������
		jf.setSize(650, 450);
		jf.setLocationRelativeTo(null); // �Ѵ���λ�����õ���Ļ����
		jf.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE); // ��������ڵĹرհ�ťʱ�˳�����û����һ�䣬���򲻻��˳���
		jf.setResizable(false);
		
		JPanel jp = new JPanel();
		
		JScrollPane jscrollpane = new JScrollPane();
		
		
		final DefaultTableModel tableModel = new DefaultTableModel();
		
		tableModel.getDataVector().clear();	//���tableModel
		
		final JTable table = new JTable(tableModel){
			public boolean isCellEditable(int row, int column){
				
				if (column != 4) {
					return false;
				}
				
				return autoCreateColumnsFromModel;
				
			}
		};
		
		Object[] columnTitle = new Object[]{"������", "�����к�", "���ϱ���","��Ʒ�ͺź�����", "�Ƿ񵼳�"};//����
		
		Object[][] rowData = new Object[aL.size()][];
		
		for(int i = 0; i < aL.size(); i++){
			
			String[] str = new String[5];
			str[0] = aL.get(i).getNc8_order_number();
			str[1] = aL.get(i).getNc8_order_line_number();
			str[2] = aL.get(i).getNc8_material_code();
			str[3] = aL.get(i).getNc8_model_no1();
			
			
			rowData[i] = new Object[]{str[0], str[1], str[2], str[3]};
		}
		
		
		tableModel.setDataVector(rowData, columnTitle);
		
		table.setModel(tableModel);
		
		table.getColumnModel().getColumn(4).setCellEditor(table.getDefaultEditor(Boolean.class));
		table.getColumnModel().getColumn(4).setCellRenderer(table.getDefaultRenderer(Boolean.class));
		
		
		
		jscrollpane.setBounds(0, 0, 650, 300);
		jscrollpane.setViewportView(table);	//������Ҫ
		
		
		JButton cancelButton = new JButton("ȡ��");
		cancelButton.setBounds(165, 360, 80, 30);
		cancelButton.setFocusPainted(false);
		JButton okButton = new JButton("ȷ��");
		okButton.setBounds(405, 360, 80, 30);
		okButton.setFocusPainted(false);
		
		jp.setLayout(null);
		// ������������뵽JFrame
		jp.add(cancelButton);
		jp.add(okButton);
		
		jp.add(jscrollpane);
		jf.setContentPane(jp);
		
		//ȡ����ť����
		cancelButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				jf.dispose();
			}
			
		});
		
		//ȷ����ť����
		okButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				for (int i = 0; i < table.getRowCount(); i++) {
					Boolean b = (Boolean) table.getValueAt(i, 4);
					if (b) {
						String nc8_order_line_number = (String) table.getValueAt(i, 1);
						result.add(nc8_order_line_number);
						System.out.println(result);
					}
					
				}
				
				List<RevBean> revLineListTest = new ArrayList<>();
				int Level = 1;
				
				TCComponent[] topComponent = Common.CommonFinder("007-����", "nc8_order_number", orderNo);
				if(topComponent!=null){
						for (int c = 0; c < topComponent.length; c++) {
							if(topComponent[c] instanceof TCComponentItemRevision){
								TCComponentItemRevision orderItemRev = (TCComponentItemRevision) topComponent[c];
								try {
									String nc8_order_line_number = orderItemRev.getProperty("nc8_order_line_number");
									TCComponent[] coms =  orderItemRev.getRelatedComponents("NC8_product_drawings_rel");
									if (result.contains(nc8_order_line_number)) {
										if(coms!=null){
											for (int p = 0; p < coms.length; p++) {
												if(coms[p] instanceof TCComponentItem){
													TCComponentItem comItem = (TCComponentItem) coms[p];
													TCComponentItemRevision rev = comItem.getLatestItemRevision();
													
													String valueCode_check = rev.getProperty("nc8_value_code");//������
													System.out.println("��name = " + rev.getProperty("object_name") + ", object_type = "+ rev.getTCProperty("object_type").getStringValue() + ", " + "������ = " + valueCode_check + "��");
													if(valueCode_check.startsWith("11") && valueCode_check.substring(6, 8).equals("00")){
														//������������ ��ͷ��λΪ11�����߰�λΪ00
														//��������
														
													}else if(valueCode_check.startsWith("11") && !valueCode_check.substring(6, 8).equals("00")){
														//��װ�������� ��ͷ��λΪ11�����߰�λ��00
														
														TCComponentBOMLine topBomline = Common.GetTopBOMLine(rev, "View", null);
														if(topBomline != null){
															//topBomline��Ϊnull��˵����Bom�ṹ���������Bom�ṹ
															traverseBom(topBomline,Level,orderItemRev,revLineListTest);
														}else{
															RevBean childRevlinestruct = new RevBean(rev, 10000,orderItemRev,"","","","","");
															revLineListTest.add(childRevlinestruct);
														}
														
													}else if(valueCode_check.startsWith("13")){
														//����������� ��ͷ��λΪ13
														
														TCComponentBOMLine topBomline = Common.GetTopBOMLine(rev, "View", null);
														if(topBomline != null){
															//topBomline��Ϊnull��˵����Bom�ṹ���������Bom�ṹ
															traverseBom(topBomline,Level,orderItemRev,revLineListTest);
														}else{
															RevBean partChildRevlinestruct = new RevBean(rev, 10000,orderItemRev,"","","","","");
															revLineListTest.add(partChildRevlinestruct);
														}
														
													}else{//������������װ�����֮���
														RevBean partChildRevlinestruct = new RevBean(rev, 10000,orderItemRev,"","","","","");
														revLineListTest.add(partChildRevlinestruct);
													}
													
												}
											}
										}
										
									}
								} catch (TCException e1) {
									// TODO Auto-generated catch block
									e1.printStackTrace();
								}
								
							}
						}
						
						try {
							TCComponentFolder getReportTemplateFolder = Common.GetReportTemplateFolder(session, "temp");// �õ�temp�ļ���
							// �����õ�temp�ļ����������е����ݼ�
								for (int i = 0; i < getReportTemplateFolder.getChildren().length; i++) {
									AIFComponentContext aifComponentContext = getReportTemplateFolder.getChildren()[i];// ��ǰ���ݼ�
									InterfaceAIFComponent component = aifComponentContext.getComponent();
									if (component instanceof TCComponentDataset) {// �ж��Ƿ������ݼ�
										String file_name = component.getProperty("object_name");// �õ���ǰ���������
										if (file_name.equals("������ϸ��.xls")) {// ƥ���ļ�����
											TCComponentDataset excleDataSet = (TCComponentDataset) component;// �õ����ݼ�
											// ���ظ����ݼ�������
											InFileName = ReportCommon.FileToLocalDir(excleDataSet, "excel", TempPath);
											if ((InFileName == null) || (InFileName.length == 0)) {
												MessageBox.post("����ģ�嵼��ʧ��", "����", 1);
												break;
											}
											// д����Ӧ���ݵ�excel�ļ���
											writeDataToExcel( revLineListTest);

											break;
										} 
									}
								}
						} catch (Exception e1) {
							e1.printStackTrace();
						}
					}
				
				jf.dispose();
				
			}
			
			
		});
		
		
		
		jf.setVisible(true);
		
	}		
		
		
}
