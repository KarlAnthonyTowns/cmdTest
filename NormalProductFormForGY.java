package com.uds.detailForm.handlers;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.WindowConstants;

import ngc_utils.ComboBoxItem;
import ngc_utils.Common;
import ngc_utils.ExportCommon;
import ngc_utils.JDBCUtils;
import ngc_utils.ReportCommon;
import ngc_utils.TcUtils;

import org.apache.axis.utils.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.StringUtil;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.jacorb.idl.runtime.int_token;

import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentPerson;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;
import com.uds.detailForm.handlers.EBOMTOPBOM.SortDataset;

/**
 * 产品明细表(工业)
 * 
 * @author tancong
 * 
 *         2019-12-17
 *         @modify  2019-12-17
 */
public class NormalProductFormForGY extends AbstractHandler {

	AbstractAIFUIApplication app = null;
	TCSession session = null;
	String[] InFileName = null;
	String TempPath = "c:\\temp\\";
	Workbook wb = null;
	// 产品明细表
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
	private List<BOMLineStruct> bomLineList = new ArrayList<>();
	TCComponentBOMLine selectedBOMLine = null;
	String whole_nc8_drawing_no = "";
	private String nc8_order_number = "";
	private String nc8_order_line_number = "";
	private String temp_nc8_order_number = "";
	private String sessionUserName = "";
	static int rowNum = 4;
	static int number = 1;
	static int productRowNum = 5;
	private String userStr = "";

	@SuppressWarnings("unchecked")
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {

		rowNum = 4;
		number = 1;
		productRowNum = 5;
		whole_nc8_drawing_no = "";

		app = AIFDesktop.getActiveDesktop().getCurrentApplication();
		session = (TCSession) app.getSession();
		String sessionName = session.getUser().toString();
		userStr = sessionName.substring(0,sessionName.indexOf("("));
		System.out.println(userStr+"===========");

		try {
			// 先获取选中工序对应的版本对象
			selComp = app.getTargetComponent();
			if (selComp instanceof TCComponentBOMLine) {
				selectedBOMLine = (TCComponentBOMLine) selComp;
				tcItemrev = selectedBOMLine.getItemRevision();
				String name_1 = tcItemrev.getProperty("object_name");
				System.out.println("【选中的name = " + name_1 + "】");

				// 判断该用户是否有权限操作
				TCComponentUser tCComponentUser = (TCComponentUser) tcItemrev.getRelatedComponent("owning_user");
				String owning_user = tCComponentUser.getUserId();
				System.out.println("所有者=======================" + owning_user);
				TCComponentUser user = session.getUser();
				String userName = user.getUserId();
				sessionUserName = userName;

				System.out.println("session.getUserName()=======================" + userName);

				if (!(owning_user.trim().equals(userName.trim()))) {
					MessageBox.post("您不是该物料所有者，没有权限操作！", "提示", MessageBox.WARNING);
					return null;

				}

				/*
				 * String first_product =
				 * tcItemrev.getProperty("nc8_firstused_products"); if
				 * ("".equals(first_product) || first_product == null) {
				 * MessageBox.post("存在零件“首次用于产品属性为空”", "提示",
				 * MessageBox.WARNING); return null; }
				 */
				/*
				 * Boolean isValid =
				 * tcItemrev.isValidPropertyName("nc8_firstused_products"); if
				 * (isValid) { String nc8_firstused_products =
				 * tcItemrev.getProperty("nc8_firstused_products");
				 * System.out.println("【nc8_firstused_products】" +
				 * nc8_firstused_products); if
				 * ("".equals(nc8_firstused_products) || nc8_firstused_products
				 * == null) { String name =
				 * tcItemrev.getProperty("object_name");
				 * System.out.println("【name = " + name + "的首次用于产品属性为空】");
				 * MessageBox.post("存在零件“首次用于产品属性为空”", "提示",
				 * MessageBox.WARNING); // return null; } }
				 */
				List<TCComponentBOMLine> list = new ArrayList<>();
				getAllChild(selectedBOMLine, list);
				/*
				 * for (int i = 0; i < list.size(); i++) { TCComponentBOMLine
				 * bomLine = list.get(i); TCComponentItemRevision
				 * bomLineItemRevision = bomLine.getItemRevision(); Boolean
				 * isValidChild =
				 * bomLineItemRevision.isValidPropertyName("nc8_firstused_products"
				 * ); if (isValidChild) { String nc8_firstused_products_child =
				 * bomLineItemRevision.getProperty("nc8_firstused_products");
				 * System.out.println("【nc8_firstused_products_child】" +
				 * nc8_firstused_products_child); if
				 * ("".equals(nc8_firstused_products_child.trim()) ||
				 * nc8_firstused_products_child == null) { String name =
				 * bomLineItemRevision.getProperty("object_name");
				 * System.out.println("【name = " + name + "的首次用于产品属性为空】");
				 * MessageBox.post("存在零件“首次用于产品属性为空”", "提示",
				 * MessageBox.WARNING); return null; } } }
				 */
				Boolean notFind = true;
				// 拿到temp文件夹
				TCComponentFolder getReportTemplateFolder = Common.GetReportTemplateFolder(session, "temp");

				// 遍历拿到temp文件夹下面所有的数据集
				for (int i = 0; i < getReportTemplateFolder.getChildren().length; i++) {

					// 当前数据集
					AIFComponentContext aifComponentContext = getReportTemplateFolder.getChildren()[i];
					InterfaceAIFComponent component = aifComponentContext.getComponent();
					// 判断是否是数据集
					if (component instanceof TCComponentDataset) {

						// 拿到当前组件的名称
						String file_name = component.getProperty("object_name");

						// 匹配文件名称
						if (file_name.equals("产品明细表.xls")) {
							notFind = false;
							// if (file_name.equals("测试")) {

							// 拿到数据集
							TCComponentDataset excleDataSet = (TCComponentDataset) component;

							System.out.println("产品明细表存在------------- ");

							// 下载该数据集到本地
							InFileName = FileToLocalDir(excleDataSet, "excel", TempPath);
							if ((InFileName == null) || (InFileName.length == 0)) {
								MessageBox.post("报表模板导出失败", "错误", 1);

								break;
							}
							// 写入相应数据到excel文件中
							writeDataToExcel(InFileName);

							break;
						} else {

						}
					}
				}

				if (notFind) {
					System.out.println("产品明细表不存在------------- ");
					MessageBox.post("产品明细表不存在，请联系管理员配置。", "错误", 1);
					return null;
				}

			} else {
				MessageBox.post("请选中BOMLine对象", "提示", MessageBox.WARNING);
				return null;
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	/*
	 * 向excel中写入数据
	 */
	private void writeDataToExcel(String[] InFileName) throws InvalidFormatException, IOException, TCException {
		List<BOMLineStruct> bomLineListTest = new ArrayList<>();
		ColletcBOMView(selectedBOMLine, 1, bomLineListTest);

		// Collections.sort(bomLineList, new SortDataset());

		System.out.println("【bomLineList】start");
		for (int i = 0; i < bomLineListTest.size(); i++) {
			BOMLineStruct bomStruct = bomLineListTest.get(i);
			TCComponentBOMLine bomLine = bomStruct.BOMLine;
			TCComponentItemRevision bomLineItemRevision = bomLine.getItemRevision();
			boolean idHasStatus = idHasStatus(bomLineItemRevision);
			if (idHasStatus) {
				MessageBox.post("BOM结构中有废弃物料，请检查", "错误", 1);
				return;
			}
			String name = bomLineItemRevision.getProperty("object_name");
			System.out.println("【index = " + i + ", name = " + name + ", level = " + bomStruct.Level + "】");

		}
		System.out.println("【bomLineList】end");

		Boolean isRoot = selectedBOMLine.isRoot();

		TCComponentBOMLine lastTopbomLine = null;

		if (isRoot) {
			lastTopbomLine = selectedBOMLine;
		} else {
			TCComponentBOMLine topbomLine = selectedBOMLine.parent();
			// 拿到顶层工艺bomline
			while (topbomLine != null) {
				lastTopbomLine = topbomLine;
				topbomLine = topbomLine.parent();
			}
		}

		if (lastTopbomLine == null) {
			MessageBox.post("该工艺未关联产品！", "错误", 1);
			return;
		}
		// 拿到顶层bomline的版本信息
		TCComponentItemRevision topBomitemRevision = lastTopbomLine.getItemRevision();

		whole_nc8_drawing_no = topBomitemRevision.getProperty("nc8_drawing_no");
		System.out.println("【整机版本的图号】" + whole_nc8_drawing_no);
		
		TCComponentItem topComponentItem = topBomitemRevision.getItem();
//		TCComponentItem selComponentItem = selectedBOMLine.getItem();

		TCComponentItemRevision orderItemRevision = null;

		// 找到该产品所属的订单（NC8_order)
		AIFComponentContext[] whereReferenced = topComponentItem.whereReferenced();
		System.out.println("whereReferenced的数量为--------------" + whereReferenced.length);
		for (int i = 0; i < whereReferenced.length; i++) {

			// 判断是否是订单类型
			InterfaceAIFComponent component = whereReferenced[i].getComponent();

			if (component instanceof TCComponentItemRevision) {

				TCComponentItemRevision orderComponentItemRevision = (TCComponentItemRevision) component;

				TCProperty tcProperty = orderComponentItemRevision.getTCProperty("object_type");

				System.out.println("whereReferenced的object_type为--------------" + tcProperty.getStringValue());

				if (tcProperty.getStringValue().equals("NC8_orderRevision")) {
					orderItemRevision = orderComponentItemRevision;
					break;
				}
			} else {
				continue;
			}
		}
//		if (orderItemRevision == null) {
//			MessageBox.post("顶层Item未关联订单", "错误", 1);
//			return;
//		}
		// TCProperty tcProperty = orderItemRevision
		// .getTCProperty("object_type");
		//
		// System.out.println("whereReferenced的object_type为--------------"
		// + tcProperty.getStringValue());

		Shell shell = new Shell();
		org.eclipse.swt.widgets.MessageBox messageBox = new org.eclipse.swt.widgets.MessageBox(shell, SWT.OK | SWT.CANCEL);
		messageBox.setText("提示");
		messageBox.setMessage("是否确定要导出EXECL BOM !");
		if (messageBox.open() == SWT.OK) {
			// writeToExcel(bomLineListTest, null);
			writeToExcel(bomLineListTest, orderItemRevision);
		}

		/*
		 * if (orderItemRevision == null) { // MessageBox.post("产品关联订单出错", "错误",
		 * 1); // return; } else {
		 * 
		 * TCProperty tcProperty = orderItemRevision
		 * .getTCProperty("object_type");	
		 * 
		 * System.out.println("whereReferenced的object_type为--------------" +
		 * tcProperty.getStringValue());
		 * 
		 * Shell shell = new Shell(); org.eclipse.swt.widgets.MessageBox
		 * messageBox = new org.eclipse.swt.widgets.MessageBox( shell, SWT.OK |
		 * SWT.CANCEL); messageBox.setText("提示");
		 * messageBox.setMessage("是否要生成选中BOM对象图号属性-A01！"); if (messageBox.open()
		 * == SWT.OK) { // writeToExcel(bomLineListTest, null);
		 * writeToExcel(bomLineListTest, orderItemRevision); } }
		 */

	}

	/**
	 * // 获取数据并写入
	 * 
	 * @param relatedComponentItemRevision
	 * @throws InvalidFormatException
	 * @throws IOException
	 * @throws TCException
	 */
	private void writeToExcel(List<BOMLineStruct> bomLineList, TCComponentItemRevision orderComponentItemRevision) throws InvalidFormatException,
			IOException, TCException {

		FileInputStream fileInputStream = new FileInputStream(InFileName[0]);
		wb = WorkbookFactory.create(fileInputStream);
		sheet1 = wb.getSheetAt(0);
		sheet2 = wb.getSheetAt(1);
		sheet3 = wb.getSheetAt(2);
		sheet4 = wb.getSheetAt(3);

		/**
		 * 产品图纸目录
		 * 整机型号为nll 
		 */
		// 产品型号/产品图号(若选中为整机，则填写整机型号，若选中不为整机，则填写图号 通过物料组判断是否为整机)
		String nc8_value_code = tcItemrev.getProperty("nc8_value_code");
		boolean isWhole = nc8_value_code.startsWith("11");
		if (isWhole) {
			System.out.println("【选中的是整机】");
			String nc8_model_no = "";
			Boolean isValid = tcItemrev.isValidPropertyName("nc8_model_no");
			if (isValid) {
				nc8_model_no = tcItemrev.getProperty("nc8_model_no");
				System.out.println("nc8_model_no--------------" + nc8_model_no);
			}else {
				System.out.println("不存在属性nc8_model_no");
			}
			//型号为空   就拿图号
			if ("".equals(nc8_model_no) || nc8_model_no == null) {
				String nc8_drawing_no = "";
				Boolean isValid2 = tcItemrev.isValidPropertyName("nc8_drawing_no");
				if (isValid2) {
					nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
					System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
				}else {
					System.out.println("不存在属性nc8_drawing_no");
				}
				ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_drawing_no);
				ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_drawing_no);
				ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_drawing_no);
				ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_drawing_no);
			}else {
				ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_model_no);
				ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_model_no);
				ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_model_no);
				ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_model_no);
			}
		} else {
			System.out.println("【选中的不是整机】");
			String nc8_drawing_no = "";
			Boolean isValid = tcItemrev.isValidPropertyName("nc8_drawing_no");
			if (isValid) {
				nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
				System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
			}else {
				System.out.println("不存在属性nc8_drawing_no");
			}
			ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_drawing_no);
			ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_drawing_no);
			ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_drawing_no);
			ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_drawing_no);
		}

		// 产品名称
		String object_name_sel = tcItemrev.getProperty("object_name");
		System.out.println("object_name--------------" + object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet1, "H2", object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet2, "B2", object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet3, "B2", object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet4, "C9", object_name_sel);
		
		
		// 产品图号
		String top_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
		System.out.println("top_drawing_no--------------" + top_drawing_no);
		ngc_utils.DoExcel.FillCell(sheet4, "C10", top_drawing_no);
		
		

		if (orderComponentItemRevision != null) {
			System.out.println("【订单】" + orderComponentItemRevision.getProperty("object_name"));
			TCComponentItem item = orderComponentItemRevision.getItem();
			TCComponentItemRevision latestItemRevision = item.getLatestItemRevision();

			// 销售订单号订单号+"-"+订单行号） 1.判断选中的条目行是不是顶层，如果不是顶层，就获取当前视图的顶层Bmline获取
			temp_nc8_order_number = latestItemRevision.getProperty("nc8_order_number");
			nc8_order_line_number = latestItemRevision.getProperty("nc8_order_line_number");
			System.out.println("订单号 nc8_order_number=" + temp_nc8_order_number);
			System.out.println("订单行号nc8_order_line_number=" + nc8_order_line_number);
			nc8_order_number = temp_nc8_order_number + "-" + nc8_order_line_number;
			System.out.println("合并之后的顶层订单号 nc8_order_number=" + nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet2, "B3", nc8_order_number);
//			ngc_utils.DoExcel.FillCell(sheet1, "M1", nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet3, "B3", nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet4, "C11", nc8_order_number);

			// 制令号 1.判断选中的条目行是不是顶层，如果不是顶层，就获取当前视图的顶层Bmline获取
			String nc8_model_no = latestItemRevision.getProperty("nc8_mo_number");
			System.out.println("nc8_mo_number--------------" + nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet2, "E3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet1, "M2", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet3, "D3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet4, "C12", nc8_model_no);

		}else{
			
			System.out.println("顶层未关联订单--------------------------------");
			
			
		}
		/*
		 * System.out.println("【订单】" +
		 * relatedComponentItemRevision.getProperty("object_name"));
		 * 
		 * // 销售订单号 String nc8_order_number = relatedComponentItemRevision
		 * .getProperty("nc8_order_number");
		 * System.out.println("nc8_order_number--------------" +
		 * nc8_order_number); ngc_utils.DoExcel.FillCell(sheet2, "B3",
		 * nc8_order_number); ngc_utils.DoExcel.FillCell(sheet1, "M1",
		 * nc8_order_number);
		 * 
		 * // 制令号 String nc8_model_no = relatedComponentItemRevision
		 * .getProperty("nc8_model_no");
		 * System.out.println("nc8_model_no--------------" + nc8_model_no);
		 * ngc_utils.DoExcel.FillCell(sheet2, "E3", nc8_model_no);
		 * ngc_utils.DoExcel.FillCell(sheet1, "M2", nc8_model_no);
		 */

		// 产品明细表

		for (int i = 0; i < bomLineList.size(); i++) {

			BOMLineStruct bomLineStruct = bomLineList.get(i);
			TCComponentBOMLine bomLine = bomLineStruct.BOMLine;
			TCComponentItemRevision bomLineRevision = bomLine.getItemRevision();
			Integer level = bomLineStruct.Level;
			String nc8_material_code_check = bomLineRevision.getProperty("nc8_material_code");
			System.out.println("【name = " + bomLineRevision.getProperty("object_name") + ", object_type = "
					+ bomLineRevision.getTCProperty("object_type").getStringValue() + ", " + "物料编码 = " + nc8_material_code_check + "】");
			if (!"".equals(nc8_material_code_check) && nc8_material_code_check != null) {
				if (nc8_material_code_check.startsWith("13")
						|| (nc8_material_code_check.startsWith("11") && !nc8_material_code_check.substring(4, 6).equals("00"))) {
					// 物料编码以13开头为零件 以11开头且五六位不为00的为部装
					String nc8_firstused_products = bomLineRevision.getProperty("nc8_firstused_products");
					System.out.println("【name = " + bomLineRevision.getProperty("object_name") + ", 首次用于产品属性值为" + nc8_firstused_products + "】");
					if ("".equals(nc8_firstused_products) || nc8_firstused_products == null) {
						TCComponentUser tCComponentUserBomLine = (TCComponentUser) bomLineRevision.getRelatedComponent("owning_user");
						String owning_user = tCComponentUserBomLine.getUserId();
						System.out.println("当前bomLine所有者=======================" + owning_user);
						if (owning_user.equals(sessionUserName)) {
							MessageBox.post("“首次用于产品”值为空，无法生成BOM！", "错误", 1);
							return;
						} else {
							fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);

						}
					} else {
//						if (nc8_firstused_products.equals(whole_nc8_drawing_no)) {
							fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);
//						} else {
//							System.out.println("“首次用于产品”属性值与图号不同");
//						}
					}
				} else {
					fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);
				}
			} else {
				fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);
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

		String saveToTCFileName = "EBOM 明细表";

		String temp_revision_id = "A";

		// 命名为：所选层的图号+版本号

		// 命名为：“EBOM”+“_”+"所选结构顶层图号"+“_”+"版本号"+"两位流水号"
		String nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
		
		String nc8_material_code = "";
		if(tcItemrev.isValidPropertyName("nc8_material_code")){
			nc8_material_code =  tcItemrev.getProperty("nc8_material_code");///物料编码
		}else if(tcItemrev.isValidPropertyName("nc8_Materialnumber")){
			nc8_material_code =  tcItemrev.getProperty("nc8_Materialnumber");//物料编码
		}
		if("".equals(nc8_material_code)){
			MessageBox.post("所选的对象没有物料编码！","提醒",MessageBox.WARNING);
		}

		if (nc8_drawing_no == null || nc8_drawing_no.trim().equals("")) {
			if (!temp_nc8_order_number.equals("") || temp_nc8_order_number != null || !nc8_order_line_number.equals("") || nc8_order_line_number != null) {
//				saveToTCFileName = "EBOM" + "_" + nc8_order_number + "_" + nc8_drawing_no + "_" + tcItemrev.getProperty("item_revision_id");
				saveToTCFileName = "EBOM" + "_" + nc8_drawing_no + "_" +nc8_material_code+"_"+ tcItemrev.getProperty("item_revision_id");
			}else {
				saveToTCFileName = "EBOM" + "_" + nc8_drawing_no + "_"  +nc8_material_code+"_"+ tcItemrev.getProperty("item_revision_id");				
			}

			temp_revision_id = tcItemrev.getProperty("item_revision_id");
			
			
			MessageBox.post("当前选中BOMLine图号为空，导出次数会显示异常！", "提示", MessageBox.WARNING);
			

		} else {

			temp_revision_id = tcItemrev.getProperty("item_revision_id") + getSequenceCode(tcItemrev.getProperty("nc8_material_code"), tcItemrev.getProperty("item_revision_id"));

			if (!temp_nc8_order_number.equals("") || temp_nc8_order_number != null || !nc8_order_line_number.equals("") || nc8_order_line_number != null) {
//				saveToTCFileName = "EBOM" + "_" + nc8_order_number + "_" + nc8_drawing_no + "_" + temp_revision_id;
				saveToTCFileName = "EBOM" +  "_" + nc8_drawing_no + "_"  +nc8_material_code+"_"+ temp_revision_id;
			}else {
				saveToTCFileName = "EBOM" + "_" + nc8_drawing_no + "_"  +nc8_material_code+"_"+ temp_revision_id;			
			}

		}

		// 顶层版本号 存储命名时候的版本号
		System.out.println("顶层版本号 =" + temp_revision_id);
		ngc_utils.DoExcel.FillCell(sheet1, "O1", temp_revision_id);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G1", temp_revision_id);
		System.out.println("生成的saveToTCFileName為------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F1", temp_revision_id);
		System.out.println("生成的saveToTCFileName為------------- " + saveToTCFileName);

		// 顶层页数（总数量除以每页的单元格行数）
		int pageBum = 1;
		if (bomLineList != null && bomLineList.size() > 0) {
			pageBum = bomLineList.size() / 34;
			if (pageBum <= 0) {
				pageBum = 1;
			}
		}

		System.out.println("顶层页数 =" + temp_revision_id);
		// sheet1
		ngc_utils.DoExcel.FillCell(sheet1, "O2", userStr);
		System.out.println("顶层页数為------------- " + saveToTCFileName);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G2", pageBum + "");
		System.out.println("sheet2顶层页数為------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F2", pageBum + "");
		System.out.println("sheet3顶层页数為------------- " + saveToTCFileName);
		
		saveToTCFileName = saveToTCFileName.replace("/", "-");
		saveToTCFileName = saveToTCFileName.replace("\\", "-");

		System.out.println("saveToTCFileName = " + saveToTCFileName);
		
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

		// 将文件写入TC挂在相应的设计文档版本下面
		String object_name = tcItemrev.getProperty("object_name");
		String nc8_drawing_no1 = tcItemrev.getProperty("nc8_drawing_no");
		String item_revision_id = tcItemrev.getProperty("item_revision_id");
		String item_id = tcItemrev.getProperty("item_id");
		//	以item_id先查询数据库是否已存在该产品或部装的设计文档对象ID
		String desdocId = getDesDocId(item_id);
		if (!"".equals(desdocId)) {		//查到了设计文档对象版本ID
			//调用零组件版本...查询系统是否存在该设计文档对象
			TCComponent[] componentzj = Common.CommonFinder("零组件版本...", "ItemID", desdocId);
			if (null != componentzj) {	//系统存在该设计文档对象
				
				TCComponentItemRevision newest = ((TCComponentItemRevision)componentzj[0]).getItem().getLatestItemRevision();
				String newestRevID = newest.getProperty("item_revision_id");
				String newestNc8_document_num2 = newest.getProperty("nc8_document_num2");
				if (newestRevID.equals(item_revision_id)) {		//设计文档版本与产品版本一致，直接拿过来用
 					//	查看设计文档版本对象是否在所选的产品的BOM伪文件夹下
					TCComponentItemRevision tccir = getItemRevision("设计文档版本", desdocId, tcItemrev, "NC8_BOM");
					if(null != tccir){	//所选产品的BOM伪文件夹下面存在设计文档版本	
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
					}else {		//所选产品的BOM伪文件夹下面存在设计文档版本,但存在于系统中,直接拿过来用
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						newest.add("IMAN_specification", datasetComponent);
						tcItemrev.add("NC8_BOM", newest);
					}
					
				}else {		//设计文档版本与产品版本不一致，预示要升版
					String name = newest.getProperty("object_name");
					String description = newest.getProperty("object_desc");
					newest = newest.saveAs(item_revision_id, name, description, false, null);//升版快乐
					newest.setProperty("nc8_document_num2",newestNc8_document_num2);
					TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
					newest.add("IMAN_specification", datasetComponent);
					tcItemrev.add("NC8_BOM", newest);
					
				}
				
			}else {		//系统不存在该设计文档对象,说明被用户给删除了，那么创建一个设计文档版本对象，其ID为desdocId
				TCComponentItem item = null;
				TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
				item = itemType.create(desdocId, item_revision_id, "NC8_design_doc", object_name+"_"+nc8_drawing_no1+"_"+nc8_material_code+"_明细表" , "", null);
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
				TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
				itemRevision.add("IMAN_specification", datasetComponent);
				tcItemrev.add("NC8_BOM", itemRevision);
				
			}
			
			
			
		}else {	//没有查到关联的设计文档对象版本ID
			
			//	调用Des doc rev...查询器查询设计文档版本
			TCComponent[] componentzj = Common.CommonFinder("Des doc rev...", "nc8_business_unit,nc8_small_class,nc8_subclass,nc8_material_code", "IBD,EBOM,EBOM" + "," + nc8_material_code);
			if (null != componentzj) {	//查到系统有设计文档版本对象
				// 判断序列号获取最新版本
				TCComponentItemRevision newest = (TCComponentItemRevision)componentzj[0];
				String sequence_id = "1";
				for (int i = 0; i < componentzj.length; i++) {
					String sequence_idTemp = componentzj[i].getProperty("sequence_id");
					if (Integer.parseInt(sequence_idTemp) > Integer.parseInt(sequence_id)) {
						sequence_id = sequence_idTemp;
						newest = (TCComponentItemRevision)componentzj[i];
					}
				}
				
 				// 首先查询BOM伪文件夹下有没有设计文档
				TCComponentItemRevision tccir = getItemRevision("设计文档版本", null, tcItemrev, "NC8_BOM");
				if (null != tccir) {	//BOM伪文件夹下面有设计文档版本
					String docNc8_material_code = tccir.getProperty("nc8_material_code");
					String docRev = tccir.getProperty("item_revision_id");
					String tccirNc8_document_num2 = tccir.getProperty("nc8_document_num2");
					if (docNc8_material_code.equals(nc8_material_code) && docRev.equals(item_revision_id)) {	//直接用这个设计文档版本对象
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
						//	把文档对象的ID与产品ID关联起来
						String desDocID = tccir.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}else if (docNc8_material_code.equals(nc8_material_code)) {		//物料编码一致，但版本不一致，说明要升版
						String name = tccir.getProperty("object_name");
						String description = tccir.getProperty("object_desc");
						tccir.saveAs(item_revision_id, name, description, false, null);//升版快乐
						tccir.setProperty("nc8_document_num2",tccirNc8_document_num2);
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
//						把文档对象的ID与产品ID关联起来
						String desDocID = tccir.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}else {
						MessageBox.post("该产品设计文档版本对象的物料编码属性与产品物料编码属性不一致！请修改设计文档版本对象的物料编码属性","错误",MessageBox.ERROR);
						return;
					}
				}else {		//BOM伪文件夹下面没有设计文档版本，但在系统中
					String newestRevID = newest.getProperty("item_revision_id");
					String newestNc8_document_num2 = newest.getProperty("nc8_document_num2");
					if (newestRevID.equals(item_revision_id)) {		//与产品版本一致，则直接用
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						newest.add("IMAN_specification", datasetComponent);
						tcItemrev.add("NC8_BOM", newest);
						// 把文档对象的ID与产品ID关联起来
						String desDocID = newest.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}else {		//与产品版本不一致，则升版
						String name = newest.getProperty("object_name");
						String description = newest.getProperty("object_desc");
						newest = newest.saveAs(item_revision_id, name, description, false, null);//升版快乐
						newest.setProperty("nc8_document_num2",newestNc8_document_num2);
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						newest.add("IMAN_specification", datasetComponent);
						tcItemrev.add("NC8_BOM", newest);
//						把文档对象的ID与产品ID关联起来
						String desDocID = newest.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}
				}
				
			}else {		//查到系统没有设计文档版本对象
				TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
				String newID = itemType.getNewID();
				TCComponentItem item = itemType.create(newID, item_revision_id, "NC8_design_doc", object_name+"_"+nc8_drawing_no1+"_"+nc8_material_code+"_明细表" , "", null);
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
				TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
//				itemRevision.add("IMAN_reference", datasetComponent);
				itemRevision.add("IMAN_specification", datasetComponent);
				tcItemrev.add("NC8_BOM", itemRevision);
				
				//	将设计文档版本对象的ID绑定在产品上
				String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
				insertDesDocId(item_id, newID, insertUser);
				
			}

			
		}
		
//		TCComponentDataset datasetComponent = ReportCommon.hasDataset(datasetType, datasetName, relationObject, relationType);
//		ReportCommon.createOrUpdateExcel(OutFileName, saveToTCFileName, itemRevision, itemRevision.getType(), true);
		System.out.println("写入TC成功--------------");

		// 打开预览
		Runtime.getRuntime().exec("cmd /c start " + OutFileName);
		MessageBox.post("报表生成完毕！", "提示", 2);

	}

	/**
	 * 下载数据集
	 * 
	 * @param componentDataset
	 *            数据集对象
	 * @param namedRefName
	 *            数据集引用
	 * @param localDir
	 *            缓存目录
	 * @return
	 */
	public synchronized static String[] FileToLocalDir(TCComponentDataset componentDataset, String namedRefName, String localDir) {
		try {
			// 获取缓存路径
			File dirObject = new File(localDir);
			if (!dirObject.exists()) {
				dirObject.mkdirs();
			}

			componentDataset = componentDataset.latest();

			// 注意：命名引用[引用名]相同的文件可能存在多个
			String namedRefFileName[] = componentDataset.getFileNames(namedRefName);
			if ((namedRefFileName == null) || (namedRefFileName.length == 0)) {
				 Common.ShowTcErrAndMsg("数据集<" + componentDataset.toString() +
				 ">没有对应的命名引用!");
				return null;
			}

			String fileDirName[] = new String[namedRefFileName.length];
			for (int i = 0; i < namedRefFileName.length; i++) {
				File tempFileObject = new File(localDir, namedRefFileName[i]);
				if (tempFileObject.exists()) {
					tempFileObject.delete();
				}
				File fileObject = componentDataset.getFile(namedRefName, namedRefFileName[i], localDir);
				fileDirName[i] = fileObject.getAbsolutePath();
			}
			return fileDirName;

		} catch (Exception e) {
			 Common.ShowTcErrAndMsg("数据集<" + componentDataset.toString() +
			 ">配置错误!");
			return null;
		}
	}

	public void getAllChild(TCComponentBOMLine bomLine, List<TCComponentBOMLine> list) throws InvalidFormatException, IOException, TCException {
		AIFComponentContext[] array = bomLine.getChildren();
		if (array == null || array.length < 1) {
			return;
		} else {
			List<TCComponentBOMLine> childBomLineList = new ArrayList<>();
			for (int i = 0; i < array.length; i++) {
				TCComponentBOMLine childBomLine = (TCComponentBOMLine) array[i].getComponent();
				childBomLineList.add(childBomLine);
			}
			for (TCComponentBOMLine tcBomLine : childBomLineList) {
				list.add(tcBomLine);
				getAllChild(tcBomLine, list);
			}
		}

	}

	boolean ColletcBOMView(TCComponentBOMLine parentLine, int Level, List<BOMLineStruct> bomLineListTest) {
		Boolean isChild = false;
		try {
			AIFComponentContext[] Cmp = parentLine.getChildren();
			if (parentLine.isRoot()) {
				isChild = false;
				BOMLineStruct bomlinestruct = new BOMLineStruct(parentLine, Level);
				Add2BOMViewList(bomlinestruct, bomLineListTest);
			} else if (parentLine.getItem() != null) {
				isChild = true;
				BOMLineStruct bomlinestruct = new BOMLineStruct(parentLine, Level);
				Add2BOMViewList(bomlinestruct, bomLineListTest);
			} else if (Cmp.length != 0) {
				int state = 0;
				for (int i = 0; i < Cmp.length; i++) {
					InterfaceAIFComponent tbl = Cmp[i].getComponent();
					if ((tbl instanceof TCComponentBOMLine)) {
						// if (tbl.getType().equals("ES4_PCB")) {
						//
						// }
						state = 1;
					}
				}
				if (state == 1) {
					BOMLineStruct bomlinestruct = new BOMLineStruct(parentLine, Level);
					Add2BOMViewList(bomlinestruct, bomLineListTest);
				}
			}
			// else if (parentLine.getItem().getType().equals("ES4_PCB")) {
			//
			// if (!parentLine.getProperty("bl_ref_designator").trim()
			// .equals("")) {
			// BOMLineStruct bomlinestruct = new BOMLineStruct(parentLine,
			// Level + 1);
			// Add2BOMViewList(bomlinestruct, this.PbomLineList);
			// }
			// }
			for (int i = 0; i < Cmp.length; i++) {
				TCComponentBOMLine BOMLine = (TCComponentBOMLine) Cmp[i].getComponent();
				// if (BOMLine.getProperty("bl_uom").equals("每个")) {
				// if (!BOMLine.getProperty("bl_quantity").equals("")) {
				// if (Integer.valueOf(BOMLine.getProperty("bl_quantity"))
				// .intValue() > 1) {
				// BOMLine.unpack();
				// }
				// }
				// }
				// 判断是否发布
				/*
				 * boolean idHasStatus =
				 * TcUtils.idHasStatus(BOMLine.getItemRevision());
				 * System.out.println("【是否发布】idHasStatus = " + idHasStatus);
				 * if(idHasStatus != false){ ColletcBOMView(BOMLine, Level + 1,
				 * bomLineListTest); }
				 */
				if (!BOMLine.isRoot()) {
					TCComponentItemRevision bRevision = BOMLine.getItemRevision();
					String bString = bRevision.getProperty("object_name");
					// 判断是否展开
					String NC8_autoExpand_true = BOMLine.getProperty("NC8_autoExpand_true");
					System.out.println("【" + bString + "不是顶层，是否展开为=" + NC8_autoExpand_true + "】");
					if (!"否".equals(NC8_autoExpand_true)) {
						ColletcBOMView(BOMLine, Level + 1, bomLineListTest);
					} else {
						BOMLineStruct bomlinestruct = new BOMLineStruct(BOMLine, Level + 1);
						Add2BOMViewList(bomlinestruct, bomLineListTest);
						checkYOrNExpand(BOMLine, Level + 1, bomLineListTest);
					}
				}

				/*
				 * if (isChild) { ColletcBOMView(BOMLine, Level + 1); }else {
				 * ColletcBOMView(BOMLine, Level); }
				 */
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return true;
	}

	void Add2BOMViewList(BOMLineStruct newBOMLine, List<BOMLineStruct> bomLineList) {
		boolean HasFind = false;
		for (int i = 0; i < bomLineList.size(); i++) {
			BOMLineStruct oldBOMLine = (BOMLineStruct) bomLineList.get(i);
			if (oldBOMLine.BOMLine == newBOMLine.BOMLine) {
				if (oldBOMLine.Level < newBOMLine.Level) {
					oldBOMLine.Level = newBOMLine.Level;
				}
				HasFind = true;
				break;
			}
		}
		if (!HasFind) {
			bomLineList.add(newBOMLine);
		}
	}

	@SuppressWarnings("rawtypes")
	class SortDataset implements Comparator {
		SortDataset() {
		}

		public int compare(Object a, Object b) {
			try {
				BOMLineStruct bom1 = (BOMLineStruct) a;
				BOMLineStruct bom2 = (BOMLineStruct) b;
				return bom2.Level - bom1.Level;
			} catch (Exception localException) {
				localException.printStackTrace();
			}
			return 0;
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
	
	
	private String getSequenceCode(String sequenceName, String revision) {

		if (sequenceName != null && sequenceName.length() != 0) {

			System.out.println("上传的sequenceName为------------- " + sequenceName);
			String sequenceCode = String.valueOf(JDBCUtils.querySequenceCode(sequenceName, revision));
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
	
	

	private void fillValue(TCComponentBOMLine bomLine, TCComponentItemRevision bomLineRevision, Sheet sheet1, Integer level, Sheet sheet2)
			throws TCException {
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
		String nc8_material_code_parent = "";
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

		String object_type = bomLineRevision.getTCProperty("object_type").getStringValue();
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
				String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("【图号】--------------" + nc8_drawing_no);
				String nc8_specification = bomLineRevision.getProperty("nc8_specification");
				System.out.println("【规格】--------------" + nc8_specification);
				daihao = nc8_drawing_no + " " + nc8_specification;
				System.out.println("【代号】--------------" + daihao);
				// 层级
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 英文名称
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					}

				} else {
					System.out.println("【没有上层零件对象】");
					nc8_material_code_parent = "";
				}

				// 材料 (属性未写明)

				/**
				 * nc8_order_number =
				 * bomLineRevision.getProperty("nc8_order_number");
				 * System.out.println("【材料】--------------" + nc8_drawing_no);
				 */

				// 数量 （（bomline属性，值为0时候显示1）
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("【数量】--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// 单重
				nc8_weight = bomLineRevision.getProperty("nc8_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);

				// 木模图号 item 属性（有就拿）
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			} else if ("NC8_CastingRevision".equals(object_type) || "NC8_ForgingsRevision".equals(object_type)
					|| "NC8_WeldingRevision".equals(object_type) || "NC8_SectionRevision".equals(object_type)) {
				/**
				 * 原材料
				 */
				System.out.println("【该对象为原材料】");
				// 代号 nc8_Standard+” ”+nc8_specification
				String nc8_Standard = bomLineRevision.getProperty("nc8_Standard");
				System.out.println("【标准】--------------" + nc8_Standard);
				String nc8_Specification = bomLineRevision.getProperty("nc8_Specification");
				System.out.println("【规格】--------------" + nc8_Specification);
				String drawing_no3 = bomLineRevision.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_Specification;
				daihao = drawing_no3;//2019/01/10更改，将nc8_Standard和nc8_Specification拼接的值改为nc8_drawing_no3的值
				System.out.println("【代号】--------------" + daihao);
				// 层级
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 英文名称
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = bomLineRevision.getProperty("nc8_Materialnumber");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("【没有上层零件对象】");
					nc8_material_code_parent = "";
				}

				// 材料
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("【数量】--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// 单重
				nc8_weight = bomLineRevision.getProperty("nc8_net_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			} else if ("NC8_AssistantMatRevision".equals(object_type)) {
				/**
				 * 辅料
				 */
				System.out.println("【该对象为辅料】");
				// 代号 nc8_Standard +” ”+ nc8_model+” ”+ nc8_Specification
				String nc8_Standard = bomLineRevision.getProperty("nc8_Standard");
				System.out.println("【标准】--------------" + nc8_Standard);
				String nc8_model = bomLineRevision.getProperty("nc8_model");
				System.out.println("【型号】--------------" + nc8_model);
				String nc8_specification = bomLineRevision.getProperty("nc8_Specification");
				System.out.println("【规格】--------------" + nc8_specification);
				String drawing_no3 = bomLineRevision.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_model + " " + nc8_specification;
				daihao = drawing_no3;//2019/01/10更改，将nc8_Standard和nc8_Specification拼接的值改为nc8_drawing_no3的值
				System.out.println("【代号】--------------" + daihao);
				// 层级
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 英文名称
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = bomLineRevision.getProperty("nc8_Materialnumber");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("【没有上层零件对象】");
					nc8_material_code_parent = "";
				}

				// 材料
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
				/**
				 * 首先获取辅料数量，辅料数量为空的时候再去获取数量
				 */
				String nc8_assist_number = bomLine.getProperty("NC8_Assist_number");
				System.out.println("【辅料数量 = " + nc8_assist_number + "】");
				if ("".equals(nc8_assist_number) || nc8_assist_number == null) {
					String bl_quantity_bak = bomLine.getProperty("bl_quantity");
					System.out.println("【数量】--------------" + bl_quantity_bak);
					if (bl_quantity_bak == null || bl_quantity_bak.length() == 0) {
						bl_quantity = "1";
					}else {
						bl_quantity = bl_quantity_bak;
					}
				}else {
					bl_quantity = nc8_assist_number;
				}
				System.out.println("【excel数量】--------------" + bl_quantity);
				

				// 单重
				nc8_weight = bomLineRevision.getProperty("nc8_net_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			} else if ("NC8_test_piecesRevision".equals(object_type)) {
				/**
				 * 试验件
				 */
				System.out.println("【该对象为试验件】");
				// 代号 nc8_drawing_no
				String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("【图号】--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("【代号】--------------" + daihao);
				// 层级
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 英文名称
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("【没有上层零件对象】");
					nc8_material_code_parent = "";
				}

				// 材料
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("【数量】--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// 单重
				nc8_weight = bomLineRevision.getProperty("nc8_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			} else if ("NC8_purchasedRevision".equals(object_type)) {
				/**
				 * 外购件
				 */
				System.out.println("【该对象为外购件】");
				// 代号 nc8_drawing_no
				String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("【图号】--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("【代号】--------------" + daihao);
				// 层级
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("【层级】" + bl_sequence_no);
				// 版本
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("【版本】--------------" + item_revision_id);
				// 中文名称
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("【中文名称】--------------" + object_name);
				// 英文名称
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("【英文名称】--------------" + nc8_part_name);
				// 物料编码
				nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
				System.out.println("【物料编码】--------------" + nc8_material_code);
				// 父类编码
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("【父类编码】--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("【没有上层零件对象】");
					nc8_material_code_parent = "";
				}

				/**
				 * 材料 nc8_material+nc8_grade+nc8_hardness_level 2018-08-21修改
				 * 1.明细表中外购件ITEM的“材料”提取ITEM属性：材质+性能等级+硬度等级；
				 * 2.明细表中外购件ITEM的“备注”提取ITEM属性：特征集+BOM备注；
				 */
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("【材质】--------------" + nc8_material);
				String nc8_grade = bomLineRevision.getProperty("nc8_grade");
				System.out.println("【性能等级】--------------" + nc8_grade);
				String nc8_hardness_level = bomLineRevision.getProperty("nc8_hardness_level");
				System.out.println("【硬度等级】--------------" + nc8_hardness_level);
				nc8_material = nc8_material + " " + nc8_grade + " " + nc8_hardness_level;
				System.out.println("【材料】--------------" + nc8_material);

				// 数量 （（bomline属性，值为0时候显示1）
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("【数量】--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}
				// 单重
				nc8_weight = bomLineRevision.getProperty("nc8_weight");
				System.out.println("【单重】--------------" + nc8_weight);
				// 备注 nc8_feature_set- nc8_grade - nc8_hardness_level
				// +NC8_BOM_remark
				String nc8_feature_set = bomLineRevision.getProperty("nc8_feature_set");
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
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("【BOM备注】--------------" + NC8_BOM_remark);
				NC8_BOM_remark = (nc8_feature_set + " " + NC8_BOM_remark).trim();
				System.out.println("【备注】--------------" + NC8_BOM_remark);
				// 木模图号
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("【木模图号】--------------" + nc8_wood_pattern);
				}

				// 质量特性

			}

		} else {
			System.out.println("【该对象为普通对象】");
			// 代号 nc8_drawing_no
			String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
			System.out.println("【图号】--------------" + nc8_drawing_no);
			daihao = nc8_drawing_no;
			System.out.println("【代号】--------------" + daihao);
			// 层级
			bl_sequence_no = bomLine.getProperty("bl_sequence_no");
			System.out.println("【层级】" + bl_sequence_no);
			// 版本
			item_revision_id = bomLineRevision.getProperty("item_revision_id");
			System.out.println("【版本】--------------" + item_revision_id);
			// 中文名称
			object_name = bomLineRevision.getProperty("object_name");
			System.out.println("【中文名称】--------------" + object_name);
			// 英文名称
			nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
			System.out.println("【英文名称】--------------" + nc8_part_name);
			// 物料编码
			nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
			System.out.println("【物料编码】--------------" + nc8_material_code);
			// 父类编码
			TCComponentBOMLine parentBomLine = bomLine.parent();
			if (parentBomLine != null) {
				TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
				if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
					System.out.println("【父类编码】--------------" + nc8_material_code_parent);
				} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
					System.out.println("【父类编码】--------------" + nc8_material_code_parent);
				}
			} else {
				System.out.println("【没有上层零件对象】");
				nc8_material_code_parent = "";
			}

			// 材料 nc8_material+nc8_grade+nc8_hardness_level
			nc8_material = bomLineRevision.getProperty("nc8_material");
			System.out.println("【材料】--------------" + nc8_material);
			// 数量 （（bomline属性，值为0时候显示1）
			bl_quantity = bomLine.getProperty("bl_quantity");
			System.out.println("【数量】--------------" + bl_quantity);
			if (bl_quantity == null || bl_quantity.length() == 0) {
				bl_quantity = "1";
			}
			// 单重
			nc8_weight = bomLineRevision.getProperty("nc8_weight");
			System.out.println("【单重】--------------" + nc8_weight);
			// 备注
			NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
			System.out.println("【备注】--------------" + NC8_BOM_remark);
			// 木模图号
			if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
				nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
				System.out.println("【木模图号】--------------" + nc8_wood_pattern);
			}

			// 质量特性

		}
		// 序号（流水号自增）
		ngc_utils.DoExcel.FillCell(sheet1, "A" + rowNum, number + "");
		String blankStr = "";

		if (level > 1) {
			for (int j = 0; j < level - 1; j++) {
				blankStr = blankStr + "  ";
			}
		}
		// bl_sequence_no = blankStr + bl_sequence_no;
		bl_sequence_no = blankStr + "L" + (level - 1);
		ngc_utils.DoExcel.FillCell(sheet1, "B" + rowNum, bl_sequence_no);
		ngc_utils.DoExcel.FillCell(sheet1, "C" + rowNum, daihao);
		ngc_utils.DoExcel.FillCell(sheet1, "D" + rowNum, item_revision_id);
		ngc_utils.DoExcel.FillCell(sheet1, "E" + rowNum, object_name);
		ngc_utils.DoExcel.FillCell(sheet1, "F" + rowNum, nc8_part_name);
		ngc_utils.DoExcel.FillCell(sheet1, "G" + rowNum, nc8_material_code);
		ngc_utils.DoExcel.FillCell(sheet1, "H" + rowNum, nc8_material_code_parent);
		ngc_utils.DoExcel.FillCell(sheet1, "I" + rowNum, nc8_material);
		ngc_utils.DoExcel.FillCell(sheet1, "J" + rowNum, bl_quantity);
		ngc_utils.DoExcel.FillCell(sheet1, "K" + rowNum, nc8_weight);
		ngc_utils.DoExcel.FillCell(sheet1, "L" + rowNum, NC8_BOM_remark);
		ngc_utils.DoExcel.FillCell(sheet1, "M" + rowNum, nc8_wood_pattern);
		ngc_utils.DoExcel.FillCell(sheet1, "N" + rowNum, "");
		// System.out.println("【最终层级】" + bl_sequence_no);
		rowNum++;
		number++;

		TCComponent[] relatedComponents = bomLineRevision.getRelatedComponents("IMAN_specification");
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
		
		String nc8_firstused_products = bomLineRevision.getProperty("nc8_firstused_products");
		if (nc8_firstused_products.equals(whole_nc8_drawing_no)) {
			boolean isRoot = bomLine.isRoot();
			
			if (isCreate || isRoot) {
				/**
				 * 产品图纸目录
				 */
				// 图号
				String nc8_drawing_no_product = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("【产品图纸目录-图号】--------------" + nc8_drawing_no_product);
				ngc_utils.DoExcel.FillCell(sheet2, "A" + productRowNum, nc8_drawing_no_product);

				// 版本
				String item_revision_id_product = bomLineRevision.getProperty("item_revision_id");
				System.out.println("【产品图纸目录-版本】--------------" + item_revision_id_product);
				ngc_utils.DoExcel.FillCell(sheet2, "B" + productRowNum, item_revision_id_product);

				// 中文名称
				String object_name_product = bomLineRevision.getProperty("object_name");
				System.out.println("【产品图纸目录-中文名称】--------------" + object_name_product);
				ngc_utils.DoExcel.FillCell(sheet2, "C" + productRowNum, object_name_product);

				// 英文名称
				String nc8_part_name_product = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("【产品图纸目录-英文名称】--------------" + nc8_part_name_product);
				ngc_utils.DoExcel.FillCell(sheet2, "D" + productRowNum, nc8_part_name_product);

				// 图幅
				String nc8_drawing_size_product = bomLineRevision.getProperty("nc8_drawing_size");
				System.out.println("【产品图纸目录-图幅】--------------" + nc8_drawing_size_product);
				ngc_utils.DoExcel.FillCell(sheet2, "E" + productRowNum, nc8_drawing_size_product);

				// 页数
				String nc8_pages_product = bomLineRevision.getProperty("nc8_pages");
				System.out.println("【产品图纸目录-页数】--------------" + nc8_pages_product);
				ngc_utils.DoExcel.FillCell(sheet2, "F" + productRowNum, nc8_pages_product);

				// 备注
				String nc8_remarks_product = bomLineRevision.getProperty("nc8_remarks");
				System.out.println("【产品图纸目录-备注】--------------" + nc8_remarks_product);
				ngc_utils.DoExcel.FillCell(sheet2, "G" + productRowNum, nc8_remarks_product);

				productRowNum++;
			}
		}
		
					
	}
	
	
	
	 public static boolean idHasStatus(TCComponent component) throws TCException{
			boolean flag =false;
		TCComponent[] components = 	component.getReferenceListProperty("release_status_list");
		System.out.println("状态2"+components.length);
		for (int i = 0; i < components.length; i++) {
			String type = components[i].getProperty("object_name");
			System.out.println("状态类型"+type);
			if("NC8_Obsolete".equals(type)){
				flag= true;
				break;
			}
		}
		return flag;
	}
	 
	public void checkYOrNExpand(TCComponentBOMLine BOMLine,int Level,  List<BOMLineStruct> bomLineListTest) throws TCException {
		AIFComponentContext[] Cmp = BOMLine.getChildren();
		for (int i = 0; i < Cmp.length; i++) {
			TCComponentBOMLine childBOMLine = (TCComponentBOMLine) Cmp[i].getComponent();
			// 判断是否展开
			String NC8_autoExpand_true = childBOMLine.getProperty("NC8_autoExpand_true");
			//if ("否".equals(NC8_autoExpand_true)) {
				//判断‘非展开导出’列的值为“是”
				String NC8_Y_or_N_Expand = childBOMLine.getProperty("NC8_Y_or_N_Expand");
				if ("是".equals(NC8_Y_or_N_Expand)) {
					ColletcBOMView(childBOMLine, Level + 1, bomLineListTest);
				}
			//}
			checkYOrNExpand(childBOMLine, Level + 1, bomLineListTest);
		}
		
	}
	
	
	
	//判断对象下是否存在特定关系的零组件版本
	public TCComponentItemRevision getItemRevision(String itemType, String desDocID, TCComponent relationObject, String relationType) {
		TCComponentItemRevision tccir = null;
		try {
			TCComponent TCComponent[] = relationObject.getRelatedComponents(relationType);
			if ((TCComponent != null) && (TCComponent.length > 0)) {
				String revision = "A";
				for (int i = 0; i < TCComponent.length; i++) {
					if (null == desDocID) {
						
						if ((TCComponent[i].getProperty("object_type").equals(itemType))
								) {
							String revision1 = TCComponent[i].getProperty("item_revision_id");
							if(revision.compareTo(revision1) == 0 || revision.compareTo(revision1) == (-1)){
								revision = revision1;
								tccir = (TCComponentItemRevision) TCComponent[i];
							}
						}
					}else {
						
						if ((TCComponent[i].getProperty("object_type").equals(itemType))
								&& (TCComponent[i].getProperty("item_id").equals(desDocID))) {
							String revision1 = TCComponent[i].getProperty("item_revision_id");
							if(revision.compareTo(revision1) == 0 || revision.compareTo(revision1) == (-1)){
								revision = revision1;
								tccir = (TCComponentItemRevision) TCComponent[i];
							}
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
	
/*	
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
*/	

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
		
		
//		String bigClass=tcItemrev.getProperty("nc8_big_class");
		System.out.println("bigClass为--------------" + bigClass);
	   	
		
		//小类
		TCProperty smallClass = tcItemrev.getTCProperty("nc8_small_class");
//		String smallClass="VB";
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
		
//		try {
			//tcItemrev.setProperty("nc8_document_num2", docuNum);//属性实际为nc8_document_num是只读状态，遂增加nc8_document_num2为与之关联并相等的属性，BMIDE里设置为隐藏可写
			//MessageBox.post("生成的文档编号为："+docuNum,"提示",MessageBox.INFORMATION);
			return docuNum;
//		} catch (Exception e1) {
//			MessageBox.post(e1.toString(),"提示",MessageBox.WARNING);
//			e1.printStackTrace();
//			return "";
//		}
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

	// 当用户创建设计文档版本时，向数据库插入产品ID与设计文档ID的对应关系
	private void insertDesDocId(String productId, String desdocId, String insertUser) {
		

		Connection conn = null;// 创建一个数据库连接
		PreparedStatement pre = null; // 创建预编译语句对象
		ResultSet result = null;
		try {

			Class.forName("oracle.jdbc.driver.OracleDriver");// 加载Oracle驱动程序
			System.out.println("开始尝试连接数据库！");
			ExportCommon ec = new ExportCommon(); // 用于获取所要连接的数据库的信息，包括数据库地址、用户名、密码
			String url = ec.getOracle_url_dev();// 数据库地址
			String user = ec.getOracle_user();// 用户名
			String password = ec.getOracle_password();// 密码
			conn = DriverManager.getConnection(url, user, password);// 获取连接
			System.out.println("连接成功！");

			// sql语句
			String sql = "insert into PLM_PRODUCT_DESDOC_RELATION (PRODUCT_ID, DESDOC_ID, INSERT_USER)" + " values(?, ?, ?)";
			pre = conn.prepareStatement(sql);// 实例化预编译语句
			pre.setString(1, productId);// 设置参数，前面的1表示参数的索引，而不是表中列名的索引
			pre.setString(2, desdocId);// 设置参数，前面的2表示参数的索引，而不是表中列名的索引
			pre.setString(3, insertUser);// 设置参数，前面的3表示参数的索引，而不是表中列名的索引

			result = pre.executeQuery();// 执行查询，注意括号中不需要再加参数

		}
		catch (Exception e1) {
			e1.printStackTrace();
		}
		finally {
			try {
				// 逐一将上面的几个对象关闭，因为不关闭的话会影响性能、并且占用资源
				// 注意关闭的顺序，最后使用的最先关闭
				if (result != null)
					result.close();
				if (pre != null)
					pre.close();
				if (conn != null)
					conn.close();
				System.out.println("数据库连接已关闭！");
			}
			catch (Exception e2) {
				e2.printStackTrace();
			}
		}
	}
	
	/*
	// 更新产品ID与设计文档ID的对应关系
	private void updateDesDocId(String productId, String desdocId, String updateUser) {
		

		Connection conn = null;// 创建一个数据库连接
		PreparedStatement pre = null; // 创建预编译语句对象
		ResultSet result = null;
		try {

			Class.forName("oracle.jdbc.driver.OracleDriver");// 加载Oracle驱动程序
			System.out.println("开始尝试连接数据库！");
			ExportCommon ec = new ExportCommon(); // 用于获取所要连接的数据库的信息，包括数据库地址、用户名、密码
			String url = ec.getOracle_url_dev();// 数据库地址
			String user = ec.getOracle_user();// 用户名
			String password = ec.getOracle_password();// 密码
			conn = DriverManager.getConnection(url, user, password);// 获取连接
			System.out.println("连接成功！");

			// sql语句
			String sql = "update PLM_PRODUCT_DESDOC_RELATION set DESDOC_ID = ?, UPDATE_USER = ? where PRODUCT_ID = ?";
			pre = conn.prepareStatement(sql);// 实例化预编译语句
			pre.setString(1, desdocId);// 设置参数，前面的1表示参数的索引，而不是表中列名的索引
			pre.setString(2, updateUser);// 设置参数，前面的2表示参数的索引，而不是表中列名的索引
			pre.setString(3, productId);// 设置参数，前面的3表示参数的索引，而不是表中列名的索引

			result = pre.executeQuery();// 执行查询，注意括号中不需要再加参数

		}
		catch (Exception e1) {
			e1.printStackTrace();
		}
		finally {
			try {
				// 逐一将上面的几个对象关闭，因为不关闭的话会影响性能、并且占用资源
				// 注意关闭的顺序，最后使用的最先关闭
				if (result != null)
					result.close();
				if (pre != null)
					pre.close();
				if (conn != null)
					conn.close();
				System.out.println("数据库连接已关闭！");
			}
			catch (Exception e2) {
				e2.printStackTrace();
			}
		}
	}
*/	
	// 查询数据库是否已有产品ID所对应的设计文档ID
	private String getDesDocId(String productId) {
		
		String resultId = "";

		Connection conn = null;// 创建一个数据库连接
		PreparedStatement pre = null; // 创建预编译语句对象
		ResultSet result = null;
		try {

			Class.forName("oracle.jdbc.driver.OracleDriver");// 加载Oracle驱动程序
			System.out.println("开始尝试连接数据库！");
			ExportCommon ec = new ExportCommon(); // 用于获取所要连接的数据库的信息，包括数据库地址、用户名、密码
			String url = ec.getOracle_url_dev();// 数据库地址
			String user = ec.getOracle_user();// 用户名
			String password = ec.getOracle_password();// 密码
			conn = DriverManager.getConnection(url, user, password);// 获取连接
			System.out.println("连接成功！");

			// sql语句
			String sql = "select DESDOC_ID from PLM_PRODUCT_DESDOC_RELATION where PRODUCT_ID = ?";
			pre = conn.prepareStatement(sql);// 实例化预编译语句
			pre.setString(1, productId);// 设置参数，前面的1表示参数的索引，而不是表中列名的索引

			result = pre.executeQuery();// 执行查询，注意括号中不需要再加参数
			// 如果查到了，下面就会返回true，否则会返回false
			if (result.next()) {
				resultId = result.getString("DESDOC_ID");
			}
			return resultId;

		}
		catch (Exception e1) {
			e1.printStackTrace();
		}
		finally {
			try {
				// 逐一将上面的几个对象关闭，因为不关闭的话会影响性能、并且占用资源
				// 注意关闭的顺序，最后使用的最先关闭
				if (result != null)
					result.close();
				if (pre != null)
					pre.close();
				if (conn != null)
					conn.close();
				System.out.println("数据库连接已关闭！");
			}
			catch (Exception e2) {
				e2.printStackTrace();
			}
		}
		return resultId;
	}
	
	
	
	//设计文档升版时写入设计文档编号,复制的com.uds.drawingNumber.manually.handlers.CopyItemRevProperty的代码
	public boolean writeIntoNumber(TCComponentItemRevision tcItemrev){
		

		try {
			String p_document_no = "";
			String p_material_code = "";
			String p_materialnumber = "";
			String p_drawing_no = "";
			String p_object_type = "";
			String p_value_code = "";
			String p_nc8_firstused_products = "";

			String p_ver = "";

			String p_sel_ver = tcItemrev.getProperty("item_revision_id");


			TCComponentItem p_item = tcItemrev.getItem();

			p_object_type = p_item.getTCProperty("object_type").getStringValue();

			TCComponent[] p_itemv = p_item.getRelatedComponents("revision_list");

			if (p_itemv.length > 0) {

				p_ver = p_itemv[0].getProperty("item_revision_id");
				// 根据物料组判断零件类型 : p_value_code
				p_value_code = p_itemv[0].getProperty("nc8_value_code");

				p_document_no = p_itemv[0].getProperty("nc8_document_num");
				p_material_code = p_itemv[0].getProperty("nc8_material_code");
				p_materialnumber = p_itemv[0].getProperty("nc8_Materialnumber");
				
				p_nc8_firstused_products = p_itemv[0].getProperty("nc8_firstused_products");
				
				if (p_object_type.equals("NC8_Casting")
						|| p_object_type.equals("NC8_Forgings")
						|| p_object_type.equals("NC8_Welding")
						|| p_object_type.equals("NC8_ProceCasting")
						|| p_object_type.equals("NC8_ProceForging")
						|| p_object_type.equals("NC8_DesignPart")) {
					p_drawing_no = p_itemv[0].getProperty("nc8_Drawing_no");
				}
				else {
					p_drawing_no = p_itemv[0].getProperty("nc8_drawing_no");
				}
			}
			
			if(p_value_code.equals("")) {
				p_value_code = "00000000";
			}
			
			// 如果选择的item类型是文档
			// 判断文档类型 : p_object_type
			// NC8_design_doc 设计文档
			// NC8_general_doc 通用文档
			// NC8_process_doc 工艺文档
			// NC8_standard_doc 标准文档
			if (p_object_type.equals("NC8_design_doc")
					|| p_object_type.equals("NC8_general_doc")
					|| p_object_type.equals("NC8_process_doc")
					|| p_object_type.equals("NC8_standard_doc")) {
				// initUI();
				tcItemrev.setProperty("nc8_document_num2",p_document_no);
				MessageBox.post("成功写入 文档编号 : " + p_document_no, "复制成功",MessageBox.INFORMATION);
			}
			
		} catch (TCException e) {
			// TODO Auto-generated catch block
			MessageBox.post(e.getDetailsMessage() + "获取先前版本的文档编号失败, 请联系管理员查询更多报错内容! ", "错误",MessageBox.ERROR);
			e.printStackTrace();
		}
		return false;
	}
	
	

}
