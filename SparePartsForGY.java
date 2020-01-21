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
 * 备件明细表
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
	// 备件明细表
	Sheet sheet1 = null;
	// 产品图纸目录
	Sheet sheet2 = null;
	// 版本变更记录
	Sheet sheet3 = null;
	// 封面
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
	       final JFrame jf = new JFrame("订单号查询");
	       jf.setSize(500, 200);
	       jf.setResizable(false);
	       jf.setLocationRelativeTo(null);
	       
	       JPanel panel = new JPanel();
	       panel.setLayout(null);
	      
	       Integer START_X=130;
	       final JLabel jl = new JLabel("请输入订单号");
	       jl.setBounds(START_X, 20, 230, 30);
	       
	       final JTextField tf=new JTextField();
	       tf.setBounds(START_X, 50, 230, 30);
	       
	       JButton btnOK = new JButton("查询");
	       btnOK.addActionListener(new ActionListener() {    
				public void actionPerformed(ActionEvent e) {
					orderNo=new StringBuffer().append(tf.getText()).toString();
					if (orderNo.length()<1) {
						JOptionPane.showMessageDialog(null, "请输入订单号后重试!");
						return;
					}
					
					
					//	通过订单号调用Packing List...查询器查询备件装箱单
					TCComponent[] packings = Common.CommonFinder("Packing List...", "nc8_Sales_Order_No", orderNo);
					if (null != packings) {
						TCComponentItemRevision packing = (TCComponentItemRevision) packings[0];
						try {
							String nc8_Line_No = packing.getProperty("nc8_Line_No");
							String[] orderLineNums = nc8_Line_No.split(",");
							
							//	给orderLineNums从小到大排个序
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
								//	通过订单号和订单行号调用007-订单查询出订单
								TCComponent[] topComponent = Common.CommonFinder("007-订单", "nc8_order_number,nc8_order_line_number", orderNo + "," + orderLineNums[j]);
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
							
							// 打开选择的窗口,让用户选择要导出哪些行
							querySparePartsUI(aL);
							
						} catch (TCException e1) {
							e1.printStackTrace();
						}
					}else{
							MessageBox.post("未找到[订单号]为"+orderNo+"的订单对象!", "提示",MessageBox.INFORMATION);
							jf.dispose();
							tf.setText("");
						}
					
					jf.dispose();
					
				}

	       });
	       btnOK.setBounds(300,120,80,30);
	       
	       JButton btnCancel = new JButton("取消");
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
		
		String sequenceNo = null;//Bom属性 层级
		TCComponentBOMLine parentBomLine = null;//Bom 父Bom
		String parentMaterialCode = null;//Bom属性 父级物料编码
		String bomQuantity = null;//Bom属性 数量
		String bomRemark = null;//Bom属性 备注
		String woodPattern = null;//Bom属性 木模图号
		
		//Bom属性 层级
		sequenceNo = topBomline.getProperty("bl_sequence_no");
		
		//Bom属性 父级物料编码
		parentBomLine = topBomline.parent();
		if (parentBomLine != null) {
			TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
			if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
				parentMaterialCode = parentItemRevision.getProperty("nc8_material_code");
			} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
				parentMaterialCode = parentItemRevision.getProperty("nc8_material_number");
			}
		} else {
			System.out.println("【没有上层零件对象】");
			parentMaterialCode = "";
		}
		
		//Bom属性 数量
		bomQuantity = topBomline.getProperty("bl_quantity");
		if (bomQuantity == null || bomQuantity.length() == 0) {
			bomQuantity = "1";
		}
		
		//Bom属性 备注
		bomRemark = topBomline.getProperty("NC8_BOM_remark");
		
		//Bom属性 木模图号
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
					// 判断是否展开
					String NC8_autoExpand_true = childBomline.getProperty("NC8_autoExpand_true");
					if (!"否".equals(NC8_autoExpand_true)) {
//						RevBean revlinestruct = new RevBean(childBomline.getItemRevision(), Level+1,orderItemRev,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
//						Add2BOMViewList(revlinestruct, bomLineListTest);
						traverseBom(childBomline, Level + 1,orderItemRev, bomLineListTest);
					} else {
						
						//Bom属性 层级
						sequenceNo = childBomline.getProperty("bl_sequence_no");
						//Bom属性 父级物料编码
						parentBomLine = childBomline.parent();
						if (parentBomLine != null) {
							TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
							if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
								parentMaterialCode = parentItemRevision.getProperty("nc8_material_code");
							} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
								parentMaterialCode = parentItemRevision.getProperty("nc8_material_number");
							}
						} else {
							System.out.println("【没有上层零件对象】");
							parentMaterialCode = "";
						}
						
						//Bom属性 数量
						bomQuantity = childBomline.getProperty("bl_quantity");
						if (bomQuantity == null || bomQuantity.length() == 0) {
							bomQuantity = "1";
						}
						
						//Bom属性 备注
						bomRemark = childBomline.getProperty("NC8_BOM_remark");
						
						//Bom属性 木模图号
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
			// 判断是否展开
			String NC8_autoExpand_true = childBOMLine.getProperty("NC8_autoExpand_true");
//			if ("否".equals(NC8_autoExpand_true)) {
				//判断‘非展开导出’列的值为“是”
				String NC8_Y_or_N_Expand = childBOMLine.getProperty("NC8_Y_or_N_Expand");
				if ("是".equals(NC8_Y_or_N_Expand)) {
					//traverseBom(childBOMLine, Level + 1,orderItemRev, bomLineListTest);
					
					
					String sequenceNo = null;//Bom属性 层级
					TCComponentBOMLine parentBomLine = null;//Bom 父Bom
					String parentMaterialCode = null;//Bom属性 父级物料编码
					String bomQuantity = null;//Bom属性 数量
					String bomRemark = null;//Bom属性 备注
					String woodPattern = null;//Bom属性 木模图号
					
					//Bom属性 层级
					sequenceNo = childBOMLine.getProperty("bl_sequence_no");
					//Bom属性 父级物料编码
					parentBomLine = childBOMLine.parent();
					if (parentBomLine != null) {
						TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
						if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
							parentMaterialCode = parentItemRevision.getProperty("nc8_material_code");
						} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
							parentMaterialCode = parentItemRevision.getProperty("nc8_material_number");
						}
					} else {
						System.out.println("【没有上层零件对象】");
						parentMaterialCode = "";
					}
					//Bom属性 数量
					bomQuantity = childBOMLine.getProperty("bl_quantity");
					if (bomQuantity == null || bomQuantity.length() == 0) {
						bomQuantity = "1";
					}
					//Bom属性 备注
					bomRemark = childBOMLine.getProperty("NC8_BOM_remark");
					//Bom属性 木模图号
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

		System.out.println("【revLineListTest】start");
		for (int i = 0; i < revLineListTest.size(); i++) {
			RevBean revStruct = revLineListTest.get(i);
			TCComponentItemRevision revLine = revStruct.RevLine;
//			TCComponentItemRevision bomLineItemRevision = bomLine.getItemRevision();
//			boolean idHasStatus = idHasStatus(bomLineItemRevision);
//			if (idHasStatus) {
//				MessageBox.post("BOM结构中有废弃物料，请检查", "错误", 1);
//				return;
//			}
			String name = revLine.getProperty("object_name");
			System.out.println("【index = " + i + ", name = " + name + ", level = " + revStruct.Level + "】");

		}
		System.out.println("【revLineListTest】end");

//		Shell shell = new Shell();
//		org.eclipse.swt.widgets.MessageBox messageBox = new org.eclipse.swt.widgets.MessageBox(shell, SWT.OK | SWT.CANCEL);
//		messageBox.setText("提示");
//		messageBox.setMessage("是否确定要导出EXECL BOM !");
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
		}catch(Exception e){System.out.println("排序错误");}
		
		/**
		 * 产品图纸目录
		 * 整机型号为nll 
		 */
		// 产品型号/产品图号(若选中为整机，则填写整机型号，若选中不为整机，则填写图号 通过物料组判断是否为整机)
//		String nc8_value_code = tcItemrev.getProperty("nc8_value_code");
//		boolean isWhole = nc8_value_code.startsWith("11");
//		if (isWhole) {
//			System.out.println("【选中的是整机】");
//			String nc8_model_no = "";
//			Boolean isValid = tcItemrev.isValidPropertyName("nc8_model_no");
//			if (isValid) {
//				nc8_model_no = tcItemrev.getProperty("nc8_model_no");
//				System.out.println("nc8_model_no--------------" + nc8_model_no);
//			}else {
//				System.out.println("不存在属性nc8_model_no");
//			}
//			//型号为空   就拿图号
//			if ("".equals(nc8_model_no) || nc8_model_no == null) {
//				String nc8_drawing_no = "";
//				Boolean isValid2 = tcItemrev.isValidPropertyName("nc8_drawing_no");
//				if (isValid2) {
//					nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
//					System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
//				}else {
//					System.out.println("不存在属性nc8_drawing_no");
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
//			System.out.println("【选中的不是整机】");
//			String nc8_drawing_no = "";
//			Boolean isValid = tcItemrev.isValidPropertyName("nc8_drawing_no");
//			if (isValid) {
//				nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
//				System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
//			}else {
//				System.out.println("不存在属性nc8_drawing_no");
//			}
//			//ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_drawing_no);
//			ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_drawing_no);
//			ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_drawing_no);
//			ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_drawing_no);
//		}
//
//		// 产品名称
//		String object_name_sel = tcItemrev.getProperty("object_name");
//		System.out.println("object_name--------------" + object_name_sel);
//		//ngc_utils.DoExcel.FillCell(sheet1, "H2", object_name_sel);
//		ngc_utils.DoExcel.FillCell(sheet2, "B2", object_name_sel);
//		ngc_utils.DoExcel.FillCell(sheet3, "B2", object_name_sel);
//		ngc_utils.DoExcel.FillCell(sheet4, "C9", object_name_sel);
//		
//		
//		// 产品图号
//		String top_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
//		System.out.println("top_drawing_no--------------" + top_drawing_no);
//		ngc_utils.DoExcel.FillCell(sheet4, "C10", top_drawing_no);
		
		
		TCComponentItemRevision orderItemRevision = null;
		if(revLineList.size()>0){
			orderItemRevision = revLineList.get(0).OrderRevLine;
		}
		 
		if (orderItemRevision != null) {
			System.out.println("【订单】" + orderItemRevision.getProperty("object_name"));
			TCComponentItem item = orderItemRevision.getItem();
			TCComponentItemRevision latestItemRevision = item.getLatestItemRevision();

			// 销售订单号
			temp_nc8_order_number = latestItemRevision.getProperty("nc8_order_number");//订单号
			System.out.println("订单号 nc8_order_number=" + temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet2, "B3", temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet1, "M1", temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet3, "B3", temp_nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet4, "C11", temp_nc8_order_number);

			// 制令号 
			String nc8_model_no = latestItemRevision.getProperty("nc8_mo_number");
			System.out.println("nc8_mo_number--------------" + nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet2, "E3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet1, "M2", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet3, "D3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet4, "C12", nc8_model_no);

		}else{
			
			System.out.println("未获取到订单--------------------------------");
			
			
		}

		// 产品明细表

		for (int i = 0; i < revLineList.size(); i++) {

			RevBean revbean = revLineList.get(i);
			TCComponentItemRevision revLine = revbean.RevLine;
			TCComponentItemRevision orderRevLine = revbean.OrderRevLine;
			String sequenceNo = revbean.sequenceNo;//Bom属性 层级
			String parentMaterialCode = revbean.parentMaterialCode;//Bom属性 父级物料编码
			String bomQuantity = revbean.bomQuantity;//Bom属性 数量
			String bomRemark = revbean.bomRemark;//Bom属性 备注
			String woodPattern = revbean.woodPattern;//Bom属性 木模图号
//			TCComponentItemRevision bomLineRevision = bomLine.getItemRevision();
			Integer level = revbean.Level;
			String nc8_material_code_check = revLine.getProperty("nc8_material_code");
			System.out.println("【name = " + revLine.getProperty("object_name") + ", object_type = "+ revLine.getTCProperty("object_type").getStringValue() + ", " + "物料编码 = " + nc8_material_code_check + "】");
			if (!"".equals(nc8_material_code_check) && nc8_material_code_check != null) {
				if (nc8_material_code_check.startsWith("13") || (nc8_material_code_check.startsWith("11") && !nc8_material_code_check.substring(4, 6).equals("00"))) {
					// 物料编码以13开头为零件 以11开头且五六位不为00的为部装
					String nc8_firstused_products = revLine.getProperty("nc8_firstused_products");
					System.out.println("【name = " + revLine.getProperty("object_name") + ", 首次用于产品属性值为" + nc8_firstused_products + "】");
					if ("".equals(nc8_firstused_products) || nc8_firstused_products == null) {
						TCComponentUser tCComponentUserBomLine = (TCComponentUser) revLine.getRelatedComponent("owning_user");
						String owning_user = tCComponentUserBomLine.getUserId();
						System.out.println("当前bomLine所有者=======================" + owning_user);
						if (owning_user.equals(session.getUser().getUserId())) {
							MessageBox.post(revLine.getProperty("object_name")+"“首次用于产品”值为空！", "错误", 1);
							return;
						} else {
							fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);

						}
					} else {
						//if (nc8_firstused_products.equals(whole_nc8_drawing_no)) {
							fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
						//} else {
							//System.out.println("“首次用于产品”属性值与图号不同");
						//}
					}
				} else {
					fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
				}
			} else {
				fillValue(revLine, orderRevLine, sheet1, level, sheet2,sequenceNo,parentMaterialCode,bomQuantity,bomRemark,woodPattern);
			}
		}
		System.out.println("写入属性值到Excel完毕------------- ");

		// 存储文件名称（当前日期）
		String time = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
		if (InFileName[0].endsWith(".xls")) {
			OutFileName = TempPath + time + ".xls";
		}
		if (InFileName[0].endsWith(".xlsx")) {
			OutFileName = TempPath + time + ".xlsx";
		}

		String saveToTCFileName = "备件明细表";

		String temp_revision_id = "A";

		// 命名为：所选层的图号+版本号

		// 命名为：“EBOM”+“_”+"所选结构顶层图号"+“_”+"版本号"+"两位流水号"
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
//			MessageBox.post("当前选中BOMLine图号为空，导出次数会显示异常！", "提示", MessageBox.WARNING);
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
		// 存储文件名称（当前日期）
		String timeNow = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
		
		saveToTCFileName = "EBOM" + "_" + orderNo + "_" +timeNow;

		// 顶层版本号 存储命名时候的版本号
		System.out.println("顶层版本号 =" + temp_revision_id);
		ngc_utils.DoExcel.FillCell(sheet1, "O1", temp_revision_id);
		//编制
		ngc_utils.DoExcel.FillCell(sheet1, "O2", userStr);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G1", temp_revision_id);
		System.out.println("生成的saveToTCFileName為------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F1", temp_revision_id);
		System.out.println("生成的saveToTCFileName為------------- " + saveToTCFileName);

		// 顶层页数（总数量除以每页的单元格行数）
		int pageBum = 1;
		if (revLineList != null && revLineList.size() > 0) {
			pageBum = revLineList.size() / 34;
			if (pageBum <= 0) {
				pageBum = 1;
			}
		}

		System.out.println("顶层页数 =" + temp_revision_id);
		// sheet1
//		ngc_utils.DoExcel.FillCell(sheet1, "O2", pageBum + "");
//		System.out.println("顶层页数為------------- " + saveToTCFileName);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G2", pageBum + "");
		System.out.println("sheet2顶层页数為------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F2", pageBum + "");
		System.out.println("sheet3顶层页数為------------- " + saveToTCFileName);

		// 存储文件名称
		if (InFileName[0].endsWith(".xls")) {
			OutFileName = TempPath + saveToTCFileName + ".xls";
		}
		if (InFileName[0].endsWith(".xlsx")) {
			OutFileName = TempPath + saveToTCFileName + ".xlsx";
		}

		// 写入文件到本地
		FileOutputStream fileOut = new FileOutputStream(OutFileName);
		wb.write(fileOut);
		fileOut.close();

		System.out.println("OutFileName--------------" + OutFileName);
		
		String queryLineNumber = "10";
		while (true) {
			TCComponent[] topComponent = Common.CommonFinder("007-订单", "nc8_order_number,nc8_order_line_number", orderNo+","+queryLineNumber);
//			TCComponent[] topComponent = Common.CommonFinder("007-订单", "nc8_order_number,nc8_order_line_number", orderNo+","+"10");
			if (null != topComponent){
				if(topComponent[0] instanceof TCComponentItemRevision){
					TCComponentItemRevision orderItemRev = (TCComponentItemRevision) topComponent[0];
					String item_revision_id = orderItemRev.getProperty("item_revision_id");
					String item_id = orderItemRev.getProperty("item_id");
					String nc8_material_code = orderItemRev.getProperty("nc8_material_code");
					if("".equals(nc8_material_code)){
						MessageBox.post("订单"+orderNo+"-"+queryLineNumber+"没有物料编码！","提醒",MessageBox.WARNING);
					}
					
					// TODO
					TCComponentItemRevision tccir = getItemRevision("设计文档版本", "EBOM" + "_" + orderNo +"明细表", orderItemRev, "NC8_other_deliverables_rel");
					if (null != tccir) {
						String NC8_design_doc_revision = tccir.getProperty("item_revision_id");
						if(!item_revision_id.equals(NC8_design_doc_revision)){
							tccir=tccir.saveAs(item_revision_id); //升版快乐
							orderItemRev.add("NC8_other_deliverables_rel", tccir);
						}
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
					}else {
						TCComponentItem item = null;
						if("A".equals(item_revision_id)){
							TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
							item = itemType.create(itemType.getNewID(),item_revision_id, "NC8_design_doc", "EBOM" + "_" + orderNo +"明细表" , "", null);
							TCComponentItemRevision itemRevision = item.getLatestItemRevision();
							itemRevision.setProperty("nc8_business_unit", "IBD");
							itemRevision.setProperty("nc8_small_class", "EBOM");
							itemRevision.setProperty("nc8_subclass", "EBOM");
							if (itemRevision.isValidPropertyName("nc8_material_code")) {	//检查合法性是因为写代码的时候还没部署这个属性
								itemRevision.setProperty("nc8_material_code", nc8_material_code);
							} 
							String nc8_document_num2 = generateNumber(itemRevision);
							if("".equals(nc8_document_num2)){
								MessageBox.post("生成的文档编号不成功！","错误",MessageBox.ERROR);
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
							//查找设计文档版本版本对象
							do{
								query_revision_id = Character.toString(new_item_revision_id);
								componentzj = Common.CommonFinder("零组件版本...", "ItemID,Revision", item_id + "," + query_revision_id);
								if(null != componentzj){
									need_update_itemrevision = getItemRevision("设计文档版本", "EBOM" + "_" + orderNo +"明细表", ((TCComponentItemRevision)componentzj[0]), "NC8_other_deliverables_rel");
									if(null != need_update_itemrevision){
										break;
									}
								}
								new_item_revision_id = (char) (new_item_revision_id - 1) ;
								
							}while("A".compareTo(Character.toString(new_item_revision_id))<=0);
							
							if(null == need_update_itemrevision){//直接创建一个与订单版本一致的itemrevision
								TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
								item = itemType.create(itemType.getNewID(),item_revision_id, "NC8_design_doc", "EBOM" + "_" + orderNo +"明细表" , "", null);
								TCComponentItemRevision itemRevision = item.getLatestItemRevision();
								itemRevision.setProperty("nc8_business_unit", "IBD");
								itemRevision.setProperty("nc8_small_class", "EBOM");
								itemRevision.setProperty("nc8_subclass", "EBOM");
								if (itemRevision.isValidPropertyName("nc8_material_code")) {	//检查合法性是因为写代码的时候还没部署这个属性
									itemRevision.setProperty("nc8_material_code", nc8_material_code);
								} 
								String nc8_document_num2 = generateNumber(itemRevision);
								if("".equals(nc8_document_num2)){
									MessageBox.post("生成的文档编号不成功！","错误",MessageBox.ERROR);
								}else{
									itemRevision.setProperty("nc8_document_num2", nc8_document_num2);
								}
								orderItemRev.add("NC8_other_deliverables_rel", itemRevision);
								TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
//							itemRevision.add("IMAN_reference", datasetComponent);
								itemRevision.add("IMAN_specification", datasetComponent);
							}else{//进行升版
								TCComponentItemRevision update_itemrevision = need_update_itemrevision.saveAs(item_revision_id); //升版快乐
								orderItemRev.add("NC8_other_deliverables_rel", update_itemrevision);
								TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
								update_itemrevision.add("IMAN_specification", datasetComponent);
							}
							
						}
					}
					
					// 将文件写入TC挂在该订单号，订单行号为10的其他交付物伪文件夹下面
//				createOrUpdateExcel(OutFileName, saveToTCFileName, orderItemRev, "NC8_other_deliverables_rel", true);
					System.out.println("写入TC成功--------------");
					MessageBox.post("生成备件明细表成功，存放于"+orderNo+"-"+queryLineNumber+"对象下面","消息",MessageBox.INFORMATION);
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
//			TCComponent[] topComp = Common.CommonFinder("007-订单", "nc8_order_number", orderNo);
//			if(topComp!=null){
//				if(topComp[0] instanceof TCComponentItemRevision){
//					TCComponentItemRevision orderItemRev = (TCComponentItemRevision) topComp[0];
//					String nc8_order_line_number = orderItemRev.getProperty("nc8_order_line_number");
//					// 将文件写入TC挂在该订单号，订单行号为nc8_order_line_number的其他交付物伪文件夹下面
//					createOrUpdateExcel(OutFileName, saveToTCFileName, orderItemRev, "NC8_other_deliverables_rel", true);
//					System.out.println("写入TC成功--------------");
//				}
//			}
//		}

		

		// 打开预览
		//Runtime.getRuntime().exec("cmd /c start " + OutFileName);
//		MessageBox.post("报表生成完毕！", "提示", 2);

	}

	private void fillValue(TCComponentItemRevision revLine, TCComponentItemRevision orderRevLine,Sheet sheet12,Integer level, Sheet sheet22,String sequenceNo,String parentMaterialCode,String bomQuantity,String bomRemark,String woodPattern) throws TCException {
		System.out.println("row = " + rowNum + ", number = " + number + ", level = " + level + ", productRowNum = " + productRowNum);
		// 代号
		String daihao = "";
		// 层级
		String bl_sequence_no = "";
		// 版本
		String item_revision_id = "";
		// 中文名称
		String object_name = "";
		// 英文名称
		String nc8_part_name = "";
		// 物料编码
		String nc8_material_code = "";
		// 父类编码
		//String nc8_material_code_parent = "";
		// 材料
		String nc8_material = "";
		// 数量
		String bl_quantity = "";
		// 单重
		String nc8_weight = "";
		// 备注
		String NC8_BOM_remark = "";
		// 木模图号
		String nc8_wood_pattern = "";
		// 订单行号
		String nc8_order_line_number = "";

		String object_type = revLine.getTCProperty("object_type").getStringValue();
		System.out.println("【object_type】" + object_type);
		if ("NC8_cust_supplyRevision".equals(object_type) || "NC8_CastingRevision".equals(object_type) || "NC8_ForgingsRevision".equals(object_type)
				|| "NC8_WeldingRevision".equals(object_type) || "NC8_SectionRevision".equals(object_type)
				|| "NC8_AssistantMatRevision".equals(object_type) || "NC8_test_piecesRevision".equals(object_type)
				|| "NC8_purchasedRevision".equals(object_type)) {

			if ("NC8_cust_supplyRevision".equals(object_type)) {
				/**
				 * 客供件
				 */
				System.out.println("【该对象为客供件】");
				// 代号 nc8_drawing_no+” ”+nc8_specification
				String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
				System.out.println("【图号】--------------" + nc8_drawing_no);
				String nc8_specification = revLine.getProperty("nc8_specification");
				System.out.println("【规格】--------------" + nc8_specification);
				String drawing_no3 = revLine.getProperty("nc8_drawing_no3");
				//daihao = nc8_drawing_no + " " + nc8_specification;
				daihao = drawing_no3;//2019/01/15更改，将nc8_Standard和nc8_Specification拼接的值改为nc8_drawing_no3的值
				System.out.println("【代号】--------------" + daihao);
				// 层级
				//bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				//System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = revLine.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 订单行号
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("【订单行号】--------------" + nc8_order_line_number);
				// 英文名称
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = revLine.getProperty("nc8_material_code");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					}
//
//				} else {
//					System.out.println("【没有上层零件对象】");
//					nc8_material_code_parent = "";
//				}

				// 材料 (属性未写明)

				/**
				 * nc8_order_number =
				 * bomLineRevision.getProperty("nc8_order_number");
				 * System.out.println("【材料】--------------" + nc8_drawing_no);
				 */

				// 数量 （（bomline属性，值为0时候显示1）
//				bl_quantity = bomLine.getProperty("bl_quantity");
//				System.out.println("【数量】--------------" + bl_quantity);
//				if (bl_quantity == null || bl_quantity.length() == 0) {
//					bl_quantity = "1";
//				}

				// 单重
				nc8_weight = revLine.getProperty("nc8_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);

				// 木模图号 item 属性（有就拿）
//				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
//					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
//					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
//				}

				// 质量特性

			} else if ("NC8_CastingRevision".equals(object_type) || "NC8_ForgingsRevision".equals(object_type)
					|| "NC8_WeldingRevision".equals(object_type) || "NC8_SectionRevision".equals(object_type)) {
				/**
				 * 原材料
				 */
				System.out.println("【该对象为原材料】");
				// 代号 nc8_Standard+” ”+nc8_specification
				String nc8_Standard = revLine.getProperty("nc8_Standard");
				System.out.println("【标准】--------------" + nc8_Standard);
				String nc8_Specification = revLine.getProperty("nc8_Specification");
				System.out.println("【规格】--------------" + nc8_Specification);
				String drawing_no3 = revLine.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_Specification;
				daihao = drawing_no3;//2019/01/15更改，将nc8_Standard和nc8_Specification拼接的值改为nc8_drawing_no3的值
				System.out.println("【代号】--------------" + daihao);
				// 层级
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = revLine.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 订单行号
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("【订单行号】--------------" + nc8_order_line_number);
				// 英文名称
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = revLine.getProperty("nc8_Materialnumber");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("【没有上层零件对象】");
//					nc8_material_code_parent = "";
//				}

				// 材料
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
//				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("【数量】--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// 单重
				nc8_weight = revLine.getProperty("nc8_net_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			} else if ("NC8_AssistantMatRevision".equals(object_type)) {
				/**
				 * 辅料
				 */
				System.out.println("【该对象为辅料】");
				// 代号 nc8_Standard +” ”+ nc8_model+” ”+ nc8_Specification
				String nc8_Standard = revLine.getProperty("nc8_Standard");
				System.out.println("【标准】--------------" + nc8_Standard);
				String nc8_model = revLine.getProperty("nc8_model");
				System.out.println("【型号】--------------" + nc8_model);
				String nc8_specification = revLine.getProperty("nc8_Specification");
				System.out.println("【规格】--------------" + nc8_specification);
				String drawing_no3 = revLine.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_model + " " + nc8_specification;
				daihao = drawing_no3;//2019/01/15更改，将nc8_Standard和nc8_Specification拼接的值改为nc8_drawing_no3的值
				System.out.println("【代号】--------------" + daihao);
				// 层级
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = revLine.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 订单行号
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("【订单行号】--------------" + nc8_order_line_number);
				// 英文名称
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = revLine.getProperty("nc8_Materialnumber");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("【没有上层零件对象】");
//					nc8_material_code_parent = "";
//				}

				// 材料
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
				/**
				 * 首先获取辅料数量，辅料数量为空的时候再去获取数量
				 */
//				String nc8_assist_number = bomLine.getProperty("NC8_Assist_number");
//				System.out.println("【辅料数量 = " + nc8_assist_number + "】");
//				if ("".equals(nc8_assist_number) || nc8_assist_number == null) {
//					String bl_quantity_bak = bomLine.getProperty("bl_quantity");
//					System.out.println("【数量】--------------" + bl_quantity_bak);
//					if (bl_quantity_bak == null || bl_quantity_bak.length() == 0) {
//						bl_quantity = "1";
//					}else {
//						bl_quantity = bl_quantity_bak;
//					}
//				}else {
//					bl_quantity = nc8_assist_number;
//				}
				System.out.println("【excel数量】--------------" + bl_quantity);
				

				// 单重
				nc8_weight = revLine.getProperty("nc8_net_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			} else if ("NC8_test_piecesRevision".equals(object_type)) {
				/**
				 * 试验件
				 */
				System.out.println("【该对象为试验件】");
				// 代号 nc8_drawing_no
				String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
				System.out.println("【图号】--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("【代号】--------------" + daihao);
				// 层级
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = revLine.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 订单行号
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("【订单行号】--------------" + nc8_order_line_number);
				// 英文名称
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = revLine.getProperty("nc8_material_code");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("【没有上层零件对象】");
//					nc8_material_code_parent = "";
//				}

				// 材料
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
//				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("【数量】--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// 单重
				nc8_weight = revLine.getProperty("nc8_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			} else if ("NC8_purchasedRevision".equals(object_type)) {
				/**
				 * 外购件
				 */
				System.out.println("【该对象为外购件】");
				// 代号 nc8_drawing_no
				String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
				System.out.println("【图号】--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("【代号】--------------" + daihao);
				// 层级
//				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = revLine.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = revLine.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 订单行号
				nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
				System.out.println("【订单行号】--------------" + nc8_order_line_number);
				// 英文名称
				nc8_part_name = revLine.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = revLine.getProperty("nc8_material_code");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
//				TCComponentBOMLine parentBomLine = bomLine.parent();
//				if (parentBomLine != null) {
//					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//					}
//				} else {
//					System.out.println("【没有上层零件对象】");
//					nc8_material_code_parent = "";
//				}

				/**
				 * 材料 nc8_material+nc8_grade+nc8_hardness_level 2018-08-21修改
				 * 1.明细表中外购件ITEM的“材料”提取ITEM属性：材质+性能等级+硬度等级；
				 * 2.明细表中外购件ITEM的“备注”提取ITEM属性：特征集+BOM备注；
				 */
				nc8_material = revLine.getProperty("nc8_material");
				System.out.println("【材质】--------------" + nc8_material);
				String nc8_grade = revLine.getProperty("nc8_grade");
				System.out.println("【性能等级】--------------" + nc8_grade);
				String nc8_hardness_level = revLine.getProperty("nc8_hardness_level");
				System.out.println("【硬度等级】--------------" + nc8_hardness_level);
				nc8_material = nc8_material + " " + nc8_grade + " " + nc8_hardness_level;
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
//				bl_quantity = bomLine.getProperty("bl_quantity");
//				System.out.println("【数量】--------------" + bl_quantity);
//				if (bl_quantity == null || bl_quantity.length() == 0) {
//					bl_quantity = "1";
//				}
				// 单重
				nc8_weight = revLine.getProperty("nc8_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注 nc8_feature_set- nc8_grade - nc8_hardness_level
				// +NC8_BOM_remark
				String nc8_feature_set = revLine.getProperty("nc8_feature_set");
				System.out.println("【特征集】--------------" + nc8_feature_set);

				/**
				 * String nc8_grade = bomLineRevision.getProperty("nc8_grade");
				 * System.out.println("【性能等级】--------------" + nc8_grade);
				 * String nc8_hardness_level =
				 * bomLineRevision.getProperty("nc8_hardness_level");
				 * System.out.println("【硬度等级】--------------" +
				 * nc8_hardness_level);
				 */

				/**
				 * 08/20变更： EBOM中外购件的备注生成原则待调整: BOm备注 更新为：特征集+
				 * BOM备注【以空格连接，当无值内容，需隐藏连接符】
				 */
//				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【BOM备注】--------------" + NC8_BOM_remark);
				NC8_BOM_remark = (nc8_feature_set + " " + NC8_BOM_remark).trim();
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			}

		} else {
			System.out.println("【该对象为普通对象】");
			// 代号 nc8_drawing_no
			String nc8_drawing_no = revLine.getProperty("nc8_drawing_no");
			System.out.println("【图号】--------------" + nc8_drawing_no);
			daihao = nc8_drawing_no;
			System.out.println("【代号】--------------" + daihao);
			// 层级
//			bl_sequence_no = bomLine.getProperty("bl_sequence_no");
//			System.out.println("【层级】" + bl_sequence_no);
			// 版本
			item_revision_id = revLine.getProperty("item_revision_id");
			System.out.println("【版本】--------------" + item_revision_id);
			// 中文名称
			object_name = revLine.getProperty("object_name");
			System.out.println("【中文名称】--------------" + object_name);
			// 订单行号
			nc8_order_line_number = orderRevLine.getProperty("nc8_order_line_number");
			System.out.println("【订单行号】--------------" + nc8_order_line_number);
			// 英文名称
			nc8_part_name = revLine.getProperty("nc8_part_name");
			System.out.println("【英文名称】--------------" + nc8_part_name);
			// 物料编码
			nc8_material_code = revLine.getProperty("nc8_material_code");
			System.out.println("【物料编码】--------------" + nc8_material_code);
			// 父类编码
//			TCComponentBOMLine parentBomLine = bomLine.parent();
//			if (parentBomLine != null) {
//				TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
//				if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
//					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
//					System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//				} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
//					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
//					System.out.println("【父类编码】--------------" + nc8_material_code_parent);
//				}
//			} else {
//				System.out.println("【没有上层零件对象】");
//				nc8_material_code_parent = "";
//			}

			// 材料 nc8_material+nc8_grade+nc8_hardness_level
			nc8_material = revLine.getProperty("nc8_material");
			System.out.println("【材料】--------------" + nc8_material);
			// 数量 （（bomline属性，值为0时候显示1）
//			bl_quantity = bomLine.getProperty("bl_quantity");
//			System.out.println("【数量】--------------" + bl_quantity);
//			if (bl_quantity == null || bl_quantity.length() == 0) {
//				bl_quantity = "1";
//			}
			// 单重
			nc8_weight = revLine.getProperty("nc8_weight");
			System.out.println("【单重】--------------" + nc8_weight);
			// 备注
//			NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
			System.out.println("【备注】--------------" + NC8_BOM_remark);
			// 木模图号
			if (revLine.isValidPropertyName("nc8_wood_pattern") && (revLine.getProperty("nc8_wood_pattern").length() != 0)) {
				nc8_wood_pattern = revLine.getProperty("nc8_wood_pattern");
				System.out.println("【木模图号】--------------" + nc8_wood_pattern);
			}

			// 质量特性

		}
		// 序号（流水号自增）
		ngc_utils.DoExcel.FillCell(sheet1, "A" + rowNum, number + "");
		String blankStr = "";

		if(level == 10000){//传过来的level为10000，该行没有Bom结构，所以不存在层级，层级为空
			ngc_utils.DoExcel.FillCell(sheet1, "B" + rowNum, "");
		}else{
			if (level > 1) {
				for (int j = 0; j < level - 1; j++) {
					blankStr = blankStr + "  ";
				}
			}
			bl_sequence_no = blankStr + "L" + (level - 1);
		}
		
		ngc_utils.DoExcel.FillCell(sheet1, "B" + rowNum, bl_sequence_no);//BomLine上取 现在留空//bl_sequence_no
		ngc_utils.DoExcel.FillCell(sheet1, "C" + rowNum, daihao);
		ngc_utils.DoExcel.FillCell(sheet1, "D" + rowNum, item_revision_id);
		ngc_utils.DoExcel.FillCell(sheet1, "E" + rowNum, object_name);
		ngc_utils.DoExcel.FillCell(sheet1, "F" + rowNum, nc8_part_name);
		ngc_utils.DoExcel.FillCell(sheet1, "G" + rowNum, nc8_material_code);
		ngc_utils.DoExcel.FillCell(sheet1, "H" + rowNum, parentMaterialCode);//BomLine的父
		ngc_utils.DoExcel.FillCell(sheet1, "I" + rowNum, nc8_material);
		ngc_utils.DoExcel.FillCell(sheet1, "J" + rowNum, bomQuantity);//BomLine上取
		ngc_utils.DoExcel.FillCell(sheet1, "K" + rowNum, nc8_weight);
		ngc_utils.DoExcel.FillCell(sheet1, "L" + rowNum, bomRemark);//BomLine上取
		ngc_utils.DoExcel.FillCell(sheet1, "M" + rowNum, woodPattern);//BomLine上取
		ngc_utils.DoExcel.FillCell(sheet1, "N" + rowNum, "");
		// System.out.println("【最终层级】" + bl_sequence_no);
		ngc_utils.DoExcel.FillCell(sheet1, "P" + rowNum, nc8_order_line_number);
		rowNum++;
		number++;

		TCComponent[] relatedComponents = revLine.getRelatedComponents("IMAN_specification");
		System.out.println("【关联的item的size】" + relatedComponents.length);
		Boolean isCreate = false;
		for (int j = 0; j < relatedComponents.length; j++) {
			TCComponent tcComponent = relatedComponents[j];
			String string = tcComponent.getProperty("object_name");
			System.out.println("【object_name】" + string);
			if (tcComponent instanceof TCComponentDataset) {
				TCComponentDataset dataset = (TCComponentDataset) tcComponent;
				String objectType = dataset.getProperty("object_type");
				System.out.println("【数据类型 object_type = " + objectType + "】");
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
			 * 产品图纸目录
			 */
			// 图号
			String nc8_drawing_no_product = revLine.getProperty("nc8_drawing_no");
			System.out.println("【产品图纸目录-图号】--------------" + nc8_drawing_no_product);
			ngc_utils.DoExcel.FillCell(sheet2, "A" + productRowNum, nc8_drawing_no_product);

			// 版本
			String item_revision_id_product = revLine.getProperty("item_revision_id");
			System.out.println("【产品图纸目录-版本】--------------" + item_revision_id_product);
			ngc_utils.DoExcel.FillCell(sheet2, "B" + productRowNum, item_revision_id_product);

			// 中文名称
			String object_name_product = revLine.getProperty("object_name");
			System.out.println("【产品图纸目录-中文名称】--------------" + object_name_product);
			ngc_utils.DoExcel.FillCell(sheet2, "C" + productRowNum, object_name_product);

			// 英文名称
			String nc8_part_name_product = revLine.getProperty("nc8_part_name");
			System.out.println("【产品图纸目录-英文名称】--------------" + nc8_part_name_product);
			ngc_utils.DoExcel.FillCell(sheet2, "D" + productRowNum, nc8_part_name_product);

			// 图幅
			String nc8_drawing_size_product = revLine.getProperty("nc8_drawing_size");
			System.out.println("【产品图纸目录-图幅】--------------" + nc8_drawing_size_product);
			ngc_utils.DoExcel.FillCell(sheet2, "E" + productRowNum, nc8_drawing_size_product);

			// 页数
			String nc8_pages_product = revLine.getProperty("nc8_pages");
			System.out.println("【产品图纸目录-页数】--------------" + nc8_pages_product);
			ngc_utils.DoExcel.FillCell(sheet2, "F" + productRowNum, nc8_pages_product);

			// 备注
			String nc8_remarks_product = revLine.getProperty("nc8_remarks");
			System.out.println("【产品图纸目录-备注】--------------" + nc8_remarks_product);
			ngc_utils.DoExcel.FillCell(sheet2, "G" + productRowNum, nc8_remarks_product);

			productRowNum++;
		}			
	}


	private String getSequenceCode(String sequenceName) {

		if (sequenceName != null && sequenceName.length() != 0) {

			System.out.println("上传的sequenceName为------------- " + sequenceName);
			String sequenceCode = String.valueOf(JDBCUtils.querySequenceCode(sequenceName));
			System.out.println("返回流水码值为------------- " + sequenceCode);

			if (sequenceCode.equals("-1")) {
				MessageBox.post("流水码获取失败!", "提示", MessageBox.WARNING);
				throw new RuntimeException("获取流水码失败");
			} else {
				while (sequenceCode.length() < 2) {
					StringBuilder sb = new StringBuilder("0");
					sb.append(sequenceCode);
					sequenceCode = sb.toString();
				}
				return sequenceCode;
			}

		} else {

			MessageBox.post("选中BOM行物料编码不能为空！", "错误", 1);
			return "";
		}

	}
	
	
	
	
	//2018/12/28日
	
	//从ngc_utils.ReportCommon中拷贝过来，因为setDatasetFileToTC方法中MessageDialog.openQuestion，不会弹出“已存在，是否覆盖时”，在本类中没生效，
	//报错为null argument:The dialog should be created in UI thread 不生效的原因可能是因为本类中已经弹了一个swing界面。
	
	//而将MessageDialog.openQuestion改成JOptionPane.showOptionDialog就可以了，但是其他报表输出功能用到这里时，就会卡死，报错为SWT UI Thread is not responding!
	
	//故 将上传Excel数据集至TC的方法拷贝了一份到本类中来，没想到更好的办法，暂时这么写
//======================================================================================================================================================================	
	/**
	 * 上传Excel数据集至TC
	 * 
	 * @param localFile 本地文件名
	 * @param datasetName 上传的数据集名称
	 * @param itemRevision 挂靠对象
	 * @param relationType 挂靠类型
	 * @param replaceAlert 是否替换
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
	
	//判断对象下是否存在特定关系的数据集
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
		 * 上传本地文件至TC 
		 * TC系统中不含本地文件的数据集
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
		 * 上传本地文件至TC 
		 * TC系统中已含有本地文件的数据集
		 * @param localFile
		 * @param datasetComponent
		 * @param datasetNamedRef
		 * @param replaceAlert
		 *        true 显示是否覆盖判断对话框
		 *        false 直接覆盖 不提示
		 *        2018/12/1 因为“已存在是否覆盖”提示框报UI线程错误，故修改提示框
		 *        2018/12/28把该方法单独拷贝到本类中来，因为别的地方用到此方法时，JOptionPane.showOptionDialog会导致卡死掉，而本类只能用JOptionPane.showOptionDialog，不能用MessageDialog.openQuestion，
		 *        可能是由于本类中在弹出“已存在，是否覆盖”之前已经有了一个Swing界面，不知道是否是这个原因
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
//					boolean confirm = MessageDialog.openQuestion(null, "确认",datasetComponent.toString() + "已经存在，是否覆盖？");
//					if (confirm) {
//						if (datasetComponent.getProperty("date_released").length() > 0) {
//							Common.ShowTcErrAndMsg("对象已发布，无法替换！");
//							return;
//						} else {
//							TCDatasetType.setFiles(datasetComponent,filePathNames, namedRefs);
//						}
//					}
					Object[] options = {"确定","取消"};
					int response=JOptionPane.showOptionDialog(null, datasetComponent.toString() + "已经存在，是否覆盖？", "选项对话框标题",JOptionPane.YES_OPTION,  JOptionPane.QUESTION_MESSAGE, null, options, options[0]);
					if(response==0){
						if (datasetComponent.getProperty("date_released").length() > 0) {
							Common.ShowTcErrAndMsg("对象已发布，无法替换！");
							return;
						} else {
							TCDatasetType.setFiles(datasetComponent,filePathNames, namedRefs);
						}
					}
				}else {
					if (datasetComponent.getProperty("date_released").length() > 0) {
						Common.ShowTcErrAndMsg("对象已发布，无法替换！");
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
		
		
		
		//判断对象下是否存在特定关系的零组件版本
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
		
		
		//从/NGC_PLM/src/com/uds/drawingNumber/codeGeneration/handlers/DocNumberGeneration.java复制的
		//------------------------------------------------------------------------------------------
			
			/**
			 * 生成文档编号
			 * 文档编号规则：业务单元-大类-小类-年号+4位流水
			 * @throws TCException
			 */
			private String generateNumber(TCComponentItemRevision tcItemrev) throws TCException{
				
				TCProperty type = tcItemrev.getTCProperty("object_type");
				
				String bigClass = "";
				//业务单元
				String bussinessUnit=tcItemrev.getProperty("nc8_business_unit");
				System.out.println("bussinessUnit为--------------" + bussinessUnit);
				if(bussinessUnit == null || bussinessUnit.equals("")){
					MessageBox.post("业务单元不能为空!","提示",MessageBox.WARNING);
					return "";
				}
				//截取业务单元前三位
				String subBussinessUnit = bussinessUnit.substring(0, 3);
				System.out.println("subBussinessUnit为--------------" + subBussinessUnit);
				
				//大类 判断类别
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
				System.out.println("bigClass为--------------" + bigClass);
			   	
				
				//小类
				TCProperty smallClass = tcItemrev.getTCProperty("nc8_small_class");
//				String smallClass="VB";
				System.out.println("smallClass为--------------" + smallClass.getStringValue());
				
				if(smallClass.getStringValue()==null ||smallClass.getStringValue().length() == 0  ){
					
					MessageBox.post("小类不能为空！","提示",MessageBox.INFORMATION);
					return "";
				}
				
				
				StringBuffer sb1=new StringBuffer();
				StringBuffer sb=new StringBuffer();
				//年号
				String yy =String.valueOf(Calendar.getInstance().get(Calendar.YEAR));
				sb1.append(subBussinessUnit.trim()).append(bigClass.trim()).append(smallClass.getStringValue().trim()).append(yy.substring(yy.length()-2));	
				
				
				String  tempCode = 	sb.append(subBussinessUnit.trim()).append("-").append(bigClass.trim()).append("-").append(smallClass.getStringValue().trim()).append("-").append(yy.substring(yy.length()-2)).toString();
				
				String sequenceCode =  getSequenceCodeII(sb1.toString()).toString().trim();
				
				String 	docuNum = tempCode+sequenceCode;
				
				System.out.println("最终生成的文档编号为--------------" + docuNum);
				
//				try {
					//tcItemrev.setProperty("nc8_document_num2", docuNum);//属性实际为nc8_document_num是只读状态，遂增加nc8_document_num2为与之关联并相等的属性，BMIDE里设置为隐藏可写
					//MessageBox.post("生成的文档编号为："+docuNum,"提示",MessageBox.INFORMATION);
					return docuNum;
//				} catch (Exception e1) {
//					MessageBox.post(e1.toString(),"提示",MessageBox.WARNING);
//					e1.printStackTrace();
//					return "";
//				}
			}

			private String getSequenceCodeII(String sequenceName) {
				System.out.println("上传的sequenceName为------------- "+sequenceName);			
				String sequenceCode = String.valueOf(JDBCUtils.querySequenceCode(sequenceName));	
				System.out.println("返回流水码值为------------- "+sequenceCode);	
				
				if(sequenceCode.equals("-1")){
					MessageBox.post("流水码获取失败!","提示",MessageBox.WARNING);
					throw new RuntimeException("获取流水码失败");
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

		
	//	供用户选择导出哪些备件		
	public void querySparePartsUI(ArrayList<SparePartsInfoBean> aL){
		
		final ArrayList<String> result = new ArrayList<String>();
		
		
		final JFrame jf = new JFrame("请选择要导出的备件"); // 创建窗口
		jf.setSize(650, 450);
		jf.setLocationRelativeTo(null); // 把窗口位置设置到屏幕中心
		jf.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE); // 当点击窗口的关闭按钮时退出程序（没有这一句，程序不会退出）
		jf.setResizable(false);
		
		JPanel jp = new JPanel();
		
		JScrollPane jscrollpane = new JScrollPane();
		
		
		final DefaultTableModel tableModel = new DefaultTableModel();
		
		tableModel.getDataVector().clear();	//清除tableModel
		
		final JTable table = new JTable(tableModel){
			public boolean isCellEditable(int row, int column){
				
				if (column != 4) {
					return false;
				}
				
				return autoCreateColumnsFromModel;
				
			}
		};
		
		Object[] columnTitle = new Object[]{"订单号", "订单行号", "物料编码","产品型号和名称", "是否导出"};//列名
		
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
		jscrollpane.setViewportView(table);	//这句很重要
		
		
		JButton cancelButton = new JButton("取消");
		cancelButton.setBounds(165, 360, 80, 30);
		cancelButton.setFocusPainted(false);
		JButton okButton = new JButton("确定");
		okButton.setBounds(405, 360, 80, 30);
		okButton.setFocusPainted(false);
		
		jp.setLayout(null);
		// 将各个组件加入到JFrame
		jp.add(cancelButton);
		jp.add(okButton);
		
		jp.add(jscrollpane);
		jf.setContentPane(jp);
		
		//取消按钮监听
		cancelButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				jf.dispose();
			}
			
		});
		
		//确定按钮监听
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
				
				TCComponent[] topComponent = Common.CommonFinder("007-订单", "nc8_order_number", orderNo);
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
													
													String valueCode_check = rev.getProperty("nc8_value_code");//物料组
													System.out.println("【name = " + rev.getProperty("object_name") + ", object_type = "+ rev.getTCProperty("object_type").getStringValue() + ", " + "物料组 = " + valueCode_check + "】");
													if(valueCode_check.startsWith("11") && valueCode_check.substring(6, 8).equals("00")){
														//整机：物料组 开头两位为11，第七八位为00
														//整机跳过
														
													}else if(valueCode_check.startsWith("11") && !valueCode_check.substring(6, 8).equals("00")){
														//部装：物料组 开头两位为11，第七八位非00
														
														TCComponentBOMLine topBomline = Common.GetTopBOMLine(rev, "View", null);
														if(topBomline != null){
															//topBomline不为null，说明有Bom结构，遂遍历该Bom结构
															traverseBom(topBomline,Level,orderItemRev,revLineListTest);
														}else{
															RevBean childRevlinestruct = new RevBean(rev, 10000,orderItemRev,"","","","","");
															revLineListTest.add(childRevlinestruct);
														}
														
													}else if(valueCode_check.startsWith("13")){
														//零件：物料组 开头两位为13
														
														TCComponentBOMLine topBomline = Common.GetTopBOMLine(rev, "View", null);
														if(topBomline != null){
															//topBomline不为null，说明有Bom结构，遂遍历该Bom结构
															traverseBom(topBomline,Level,orderItemRev,revLineListTest);
														}else{
															RevBean partChildRevlinestruct = new RevBean(rev, 10000,orderItemRev,"","","","","");
															revLineListTest.add(partChildRevlinestruct);
														}
														
													}else{//除了整机，部装，零件之外的
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
							TCComponentFolder getReportTemplateFolder = Common.GetReportTemplateFolder(session, "temp");// 拿到temp文件夹
							// 遍历拿到temp文件夹下面所有的数据集
								for (int i = 0; i < getReportTemplateFolder.getChildren().length; i++) {
									AIFComponentContext aifComponentContext = getReportTemplateFolder.getChildren()[i];// 当前数据集
									InterfaceAIFComponent component = aifComponentContext.getComponent();
									if (component instanceof TCComponentDataset) {// 判断是否是数据集
										String file_name = component.getProperty("object_name");// 拿到当前组件的名称
										if (file_name.equals("备件明细表.xls")) {// 匹配文件名称
											TCComponentDataset excleDataSet = (TCComponentDataset) component;// 拿到数据集
											// 下载该数据集到本地
											InFileName = ReportCommon.FileToLocalDir(excleDataSet, "excel", TempPath);
											if ((InFileName == null) || (InFileName.length == 0)) {
												MessageBox.post("报表模板导出失败", "错误", 1);
												break;
											}
											// 写入相应数据到excel文件中
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
