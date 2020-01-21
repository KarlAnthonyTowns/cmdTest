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
 * ��Ʒ��ϸ��(��ҵ)
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
	// ��Ʒ��ϸ��
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
			// �Ȼ�ȡѡ�й����Ӧ�İ汾����
			selComp = app.getTargetComponent();
			if (selComp instanceof TCComponentBOMLine) {
				selectedBOMLine = (TCComponentBOMLine) selComp;
				tcItemrev = selectedBOMLine.getItemRevision();
				String name_1 = tcItemrev.getProperty("object_name");
				System.out.println("��ѡ�е�name = " + name_1 + "��");

				// �жϸ��û��Ƿ���Ȩ�޲���
				TCComponentUser tCComponentUser = (TCComponentUser) tcItemrev.getRelatedComponent("owning_user");
				String owning_user = tCComponentUser.getUserId();
				System.out.println("������=======================" + owning_user);
				TCComponentUser user = session.getUser();
				String userName = user.getUserId();
				sessionUserName = userName;

				System.out.println("session.getUserName()=======================" + userName);

				if (!(owning_user.trim().equals(userName.trim()))) {
					MessageBox.post("�����Ǹ����������ߣ�û��Ȩ�޲�����", "��ʾ", MessageBox.WARNING);
					return null;

				}

				/*
				 * String first_product =
				 * tcItemrev.getProperty("nc8_firstused_products"); if
				 * ("".equals(first_product) || first_product == null) {
				 * MessageBox.post("����������״����ڲ�Ʒ����Ϊ�ա�", "��ʾ",
				 * MessageBox.WARNING); return null; }
				 */
				/*
				 * Boolean isValid =
				 * tcItemrev.isValidPropertyName("nc8_firstused_products"); if
				 * (isValid) { String nc8_firstused_products =
				 * tcItemrev.getProperty("nc8_firstused_products");
				 * System.out.println("��nc8_firstused_products��" +
				 * nc8_firstused_products); if
				 * ("".equals(nc8_firstused_products) || nc8_firstused_products
				 * == null) { String name =
				 * tcItemrev.getProperty("object_name");
				 * System.out.println("��name = " + name + "���״����ڲ�Ʒ����Ϊ�ա�");
				 * MessageBox.post("����������״����ڲ�Ʒ����Ϊ�ա�", "��ʾ",
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
				 * System.out.println("��nc8_firstused_products_child��" +
				 * nc8_firstused_products_child); if
				 * ("".equals(nc8_firstused_products_child.trim()) ||
				 * nc8_firstused_products_child == null) { String name =
				 * bomLineItemRevision.getProperty("object_name");
				 * System.out.println("��name = " + name + "���״����ڲ�Ʒ����Ϊ�ա�");
				 * MessageBox.post("����������״����ڲ�Ʒ����Ϊ�ա�", "��ʾ",
				 * MessageBox.WARNING); return null; } } }
				 */
				Boolean notFind = true;
				// �õ�temp�ļ���
				TCComponentFolder getReportTemplateFolder = Common.GetReportTemplateFolder(session, "temp");

				// �����õ�temp�ļ����������е����ݼ�
				for (int i = 0; i < getReportTemplateFolder.getChildren().length; i++) {

					// ��ǰ���ݼ�
					AIFComponentContext aifComponentContext = getReportTemplateFolder.getChildren()[i];
					InterfaceAIFComponent component = aifComponentContext.getComponent();
					// �ж��Ƿ������ݼ�
					if (component instanceof TCComponentDataset) {

						// �õ���ǰ���������
						String file_name = component.getProperty("object_name");

						// ƥ���ļ�����
						if (file_name.equals("��Ʒ��ϸ��.xls")) {
							notFind = false;
							// if (file_name.equals("����")) {

							// �õ����ݼ�
							TCComponentDataset excleDataSet = (TCComponentDataset) component;

							System.out.println("��Ʒ��ϸ�����------------- ");

							// ���ظ����ݼ�������
							InFileName = FileToLocalDir(excleDataSet, "excel", TempPath);
							if ((InFileName == null) || (InFileName.length == 0)) {
								MessageBox.post("����ģ�嵼��ʧ��", "����", 1);

								break;
							}
							// д����Ӧ���ݵ�excel�ļ���
							writeDataToExcel(InFileName);

							break;
						} else {

						}
					}
				}

				if (notFind) {
					System.out.println("��Ʒ��ϸ������------------- ");
					MessageBox.post("��Ʒ��ϸ�����ڣ�����ϵ����Ա���á�", "����", 1);
					return null;
				}

			} else {
				MessageBox.post("��ѡ��BOMLine����", "��ʾ", MessageBox.WARNING);
				return null;
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	/*
	 * ��excel��д������
	 */
	private void writeDataToExcel(String[] InFileName) throws InvalidFormatException, IOException, TCException {
		List<BOMLineStruct> bomLineListTest = new ArrayList<>();
		ColletcBOMView(selectedBOMLine, 1, bomLineListTest);

		// Collections.sort(bomLineList, new SortDataset());

		System.out.println("��bomLineList��start");
		for (int i = 0; i < bomLineListTest.size(); i++) {
			BOMLineStruct bomStruct = bomLineListTest.get(i);
			TCComponentBOMLine bomLine = bomStruct.BOMLine;
			TCComponentItemRevision bomLineItemRevision = bomLine.getItemRevision();
			boolean idHasStatus = idHasStatus(bomLineItemRevision);
			if (idHasStatus) {
				MessageBox.post("BOM�ṹ���з������ϣ�����", "����", 1);
				return;
			}
			String name = bomLineItemRevision.getProperty("object_name");
			System.out.println("��index = " + i + ", name = " + name + ", level = " + bomStruct.Level + "��");

		}
		System.out.println("��bomLineList��end");

		Boolean isRoot = selectedBOMLine.isRoot();

		TCComponentBOMLine lastTopbomLine = null;

		if (isRoot) {
			lastTopbomLine = selectedBOMLine;
		} else {
			TCComponentBOMLine topbomLine = selectedBOMLine.parent();
			// �õ����㹤��bomline
			while (topbomLine != null) {
				lastTopbomLine = topbomLine;
				topbomLine = topbomLine.parent();
			}
		}

		if (lastTopbomLine == null) {
			MessageBox.post("�ù���δ������Ʒ��", "����", 1);
			return;
		}
		// �õ�����bomline�İ汾��Ϣ
		TCComponentItemRevision topBomitemRevision = lastTopbomLine.getItemRevision();

		whole_nc8_drawing_no = topBomitemRevision.getProperty("nc8_drawing_no");
		System.out.println("�������汾��ͼ�š�" + whole_nc8_drawing_no);
		
		TCComponentItem topComponentItem = topBomitemRevision.getItem();
//		TCComponentItem selComponentItem = selectedBOMLine.getItem();

		TCComponentItemRevision orderItemRevision = null;

		// �ҵ��ò�Ʒ�����Ķ�����NC8_order)
		AIFComponentContext[] whereReferenced = topComponentItem.whereReferenced();
		System.out.println("whereReferenced������Ϊ--------------" + whereReferenced.length);
		for (int i = 0; i < whereReferenced.length; i++) {

			// �ж��Ƿ��Ƕ�������
			InterfaceAIFComponent component = whereReferenced[i].getComponent();

			if (component instanceof TCComponentItemRevision) {

				TCComponentItemRevision orderComponentItemRevision = (TCComponentItemRevision) component;

				TCProperty tcProperty = orderComponentItemRevision.getTCProperty("object_type");

				System.out.println("whereReferenced��object_typeΪ--------------" + tcProperty.getStringValue());

				if (tcProperty.getStringValue().equals("NC8_orderRevision")) {
					orderItemRevision = orderComponentItemRevision;
					break;
				}
			} else {
				continue;
			}
		}
//		if (orderItemRevision == null) {
//			MessageBox.post("����Itemδ��������", "����", 1);
//			return;
//		}
		// TCProperty tcProperty = orderItemRevision
		// .getTCProperty("object_type");
		//
		// System.out.println("whereReferenced��object_typeΪ--------------"
		// + tcProperty.getStringValue());

		Shell shell = new Shell();
		org.eclipse.swt.widgets.MessageBox messageBox = new org.eclipse.swt.widgets.MessageBox(shell, SWT.OK | SWT.CANCEL);
		messageBox.setText("��ʾ");
		messageBox.setMessage("�Ƿ�ȷ��Ҫ����EXECL BOM !");
		if (messageBox.open() == SWT.OK) {
			// writeToExcel(bomLineListTest, null);
			writeToExcel(bomLineListTest, orderItemRevision);
		}

		/*
		 * if (orderItemRevision == null) { // MessageBox.post("��Ʒ������������", "����",
		 * 1); // return; } else {
		 * 
		 * TCProperty tcProperty = orderItemRevision
		 * .getTCProperty("object_type");	
		 * 
		 * System.out.println("whereReferenced��object_typeΪ--------------" +
		 * tcProperty.getStringValue());
		 * 
		 * Shell shell = new Shell(); org.eclipse.swt.widgets.MessageBox
		 * messageBox = new org.eclipse.swt.widgets.MessageBox( shell, SWT.OK |
		 * SWT.CANCEL); messageBox.setText("��ʾ");
		 * messageBox.setMessage("�Ƿ�Ҫ����ѡ��BOM����ͼ������-A01��"); if (messageBox.open()
		 * == SWT.OK) { // writeToExcel(bomLineListTest, null);
		 * writeToExcel(bomLineListTest, orderItemRevision); } }
		 */

	}

	/**
	 * // ��ȡ���ݲ�д��
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
		 * ��ƷͼֽĿ¼
		 * �����ͺ�Ϊnll 
		 */
		// ��Ʒ�ͺ�/��Ʒͼ��(��ѡ��Ϊ����������д�����ͺţ���ѡ�в�Ϊ����������дͼ�� ͨ���������ж��Ƿ�Ϊ����)
		String nc8_value_code = tcItemrev.getProperty("nc8_value_code");
		boolean isWhole = nc8_value_code.startsWith("11");
		if (isWhole) {
			System.out.println("��ѡ�е���������");
			String nc8_model_no = "";
			Boolean isValid = tcItemrev.isValidPropertyName("nc8_model_no");
			if (isValid) {
				nc8_model_no = tcItemrev.getProperty("nc8_model_no");
				System.out.println("nc8_model_no--------------" + nc8_model_no);
			}else {
				System.out.println("����������nc8_model_no");
			}
			//�ͺ�Ϊ��   ����ͼ��
			if ("".equals(nc8_model_no) || nc8_model_no == null) {
				String nc8_drawing_no = "";
				Boolean isValid2 = tcItemrev.isValidPropertyName("nc8_drawing_no");
				if (isValid2) {
					nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
					System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
				}else {
					System.out.println("����������nc8_drawing_no");
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
			System.out.println("��ѡ�еĲ���������");
			String nc8_drawing_no = "";
			Boolean isValid = tcItemrev.isValidPropertyName("nc8_drawing_no");
			if (isValid) {
				nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
				System.out.println("nc8_drawing_no--------------" + nc8_drawing_no);
			}else {
				System.out.println("����������nc8_drawing_no");
			}
			ngc_utils.DoExcel.FillCell(sheet1, "H1", nc8_drawing_no);
			ngc_utils.DoExcel.FillCell(sheet2, "B1", nc8_drawing_no);
			ngc_utils.DoExcel.FillCell(sheet3, "B1", nc8_drawing_no);
			ngc_utils.DoExcel.FillCell(sheet4, "C8", nc8_drawing_no);
		}

		// ��Ʒ����
		String object_name_sel = tcItemrev.getProperty("object_name");
		System.out.println("object_name--------------" + object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet1, "H2", object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet2, "B2", object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet3, "B2", object_name_sel);
		ngc_utils.DoExcel.FillCell(sheet4, "C9", object_name_sel);
		
		
		// ��Ʒͼ��
		String top_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
		System.out.println("top_drawing_no--------------" + top_drawing_no);
		ngc_utils.DoExcel.FillCell(sheet4, "C10", top_drawing_no);
		
		

		if (orderComponentItemRevision != null) {
			System.out.println("��������" + orderComponentItemRevision.getProperty("object_name"));
			TCComponentItem item = orderComponentItemRevision.getItem();
			TCComponentItemRevision latestItemRevision = item.getLatestItemRevision();

			// ���۶����Ŷ�����+"-"+�����кţ� 1.�ж�ѡ�е���Ŀ���ǲ��Ƕ��㣬������Ƕ��㣬�ͻ�ȡ��ǰ��ͼ�Ķ���Bmline��ȡ
			temp_nc8_order_number = latestItemRevision.getProperty("nc8_order_number");
			nc8_order_line_number = latestItemRevision.getProperty("nc8_order_line_number");
			System.out.println("������ nc8_order_number=" + temp_nc8_order_number);
			System.out.println("�����к�nc8_order_line_number=" + nc8_order_line_number);
			nc8_order_number = temp_nc8_order_number + "-" + nc8_order_line_number;
			System.out.println("�ϲ�֮��Ķ��㶩���� nc8_order_number=" + nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet2, "B3", nc8_order_number);
//			ngc_utils.DoExcel.FillCell(sheet1, "M1", nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet3, "B3", nc8_order_number);
			ngc_utils.DoExcel.FillCell(sheet4, "C11", nc8_order_number);

			// ����� 1.�ж�ѡ�е���Ŀ���ǲ��Ƕ��㣬������Ƕ��㣬�ͻ�ȡ��ǰ��ͼ�Ķ���Bmline��ȡ
			String nc8_model_no = latestItemRevision.getProperty("nc8_mo_number");
			System.out.println("nc8_mo_number--------------" + nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet2, "E3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet1, "M2", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet3, "D3", nc8_model_no);
			ngc_utils.DoExcel.FillCell(sheet4, "C12", nc8_model_no);

		}else{
			
			System.out.println("����δ��������--------------------------------");
			
			
		}
		/*
		 * System.out.println("��������" +
		 * relatedComponentItemRevision.getProperty("object_name"));
		 * 
		 * // ���۶����� String nc8_order_number = relatedComponentItemRevision
		 * .getProperty("nc8_order_number");
		 * System.out.println("nc8_order_number--------------" +
		 * nc8_order_number); ngc_utils.DoExcel.FillCell(sheet2, "B3",
		 * nc8_order_number); ngc_utils.DoExcel.FillCell(sheet1, "M1",
		 * nc8_order_number);
		 * 
		 * // ����� String nc8_model_no = relatedComponentItemRevision
		 * .getProperty("nc8_model_no");
		 * System.out.println("nc8_model_no--------------" + nc8_model_no);
		 * ngc_utils.DoExcel.FillCell(sheet2, "E3", nc8_model_no);
		 * ngc_utils.DoExcel.FillCell(sheet1, "M2", nc8_model_no);
		 */

		// ��Ʒ��ϸ��

		for (int i = 0; i < bomLineList.size(); i++) {

			BOMLineStruct bomLineStruct = bomLineList.get(i);
			TCComponentBOMLine bomLine = bomLineStruct.BOMLine;
			TCComponentItemRevision bomLineRevision = bomLine.getItemRevision();
			Integer level = bomLineStruct.Level;
			String nc8_material_code_check = bomLineRevision.getProperty("nc8_material_code");
			System.out.println("��name = " + bomLineRevision.getProperty("object_name") + ", object_type = "
					+ bomLineRevision.getTCProperty("object_type").getStringValue() + ", " + "���ϱ��� = " + nc8_material_code_check + "��");
			if (!"".equals(nc8_material_code_check) && nc8_material_code_check != null) {
				if (nc8_material_code_check.startsWith("13")
						|| (nc8_material_code_check.startsWith("11") && !nc8_material_code_check.substring(4, 6).equals("00"))) {
					// ���ϱ�����13��ͷΪ��� ��11��ͷ������λ��Ϊ00��Ϊ��װ
					String nc8_firstused_products = bomLineRevision.getProperty("nc8_firstused_products");
					System.out.println("��name = " + bomLineRevision.getProperty("object_name") + ", �״����ڲ�Ʒ����ֵΪ" + nc8_firstused_products + "��");
					if ("".equals(nc8_firstused_products) || nc8_firstused_products == null) {
						TCComponentUser tCComponentUserBomLine = (TCComponentUser) bomLineRevision.getRelatedComponent("owning_user");
						String owning_user = tCComponentUserBomLine.getUserId();
						System.out.println("��ǰbomLine������=======================" + owning_user);
						if (owning_user.equals(sessionUserName)) {
							MessageBox.post("���״����ڲ�Ʒ��ֵΪ�գ��޷�����BOM��", "����", 1);
							return;
						} else {
							fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);

						}
					} else {
//						if (nc8_firstused_products.equals(whole_nc8_drawing_no)) {
							fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);
//						} else {
//							System.out.println("���״����ڲ�Ʒ������ֵ��ͼ�Ų�ͬ");
//						}
					}
				} else {
					fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);
				}
			} else {
				fillValue(bomLine, bomLineRevision, sheet1, level, sheet2);
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

		String saveToTCFileName = "EBOM ��ϸ��";

		String temp_revision_id = "A";

		// ����Ϊ����ѡ���ͼ��+�汾��

		// ����Ϊ����EBOM��+��_��+"��ѡ�ṹ����ͼ��"+��_��+"�汾��"+"��λ��ˮ��"
		String nc8_drawing_no = tcItemrev.getProperty("nc8_drawing_no");
		
		String nc8_material_code = "";
		if(tcItemrev.isValidPropertyName("nc8_material_code")){
			nc8_material_code =  tcItemrev.getProperty("nc8_material_code");///���ϱ���
		}else if(tcItemrev.isValidPropertyName("nc8_Materialnumber")){
			nc8_material_code =  tcItemrev.getProperty("nc8_Materialnumber");//���ϱ���
		}
		if("".equals(nc8_material_code)){
			MessageBox.post("��ѡ�Ķ���û�����ϱ��룡","����",MessageBox.WARNING);
		}

		if (nc8_drawing_no == null || nc8_drawing_no.trim().equals("")) {
			if (!temp_nc8_order_number.equals("") || temp_nc8_order_number != null || !nc8_order_line_number.equals("") || nc8_order_line_number != null) {
//				saveToTCFileName = "EBOM" + "_" + nc8_order_number + "_" + nc8_drawing_no + "_" + tcItemrev.getProperty("item_revision_id");
				saveToTCFileName = "EBOM" + "_" + nc8_drawing_no + "_" +nc8_material_code+"_"+ tcItemrev.getProperty("item_revision_id");
			}else {
				saveToTCFileName = "EBOM" + "_" + nc8_drawing_no + "_"  +nc8_material_code+"_"+ tcItemrev.getProperty("item_revision_id");				
			}

			temp_revision_id = tcItemrev.getProperty("item_revision_id");
			
			
			MessageBox.post("��ǰѡ��BOMLineͼ��Ϊ�գ�������������ʾ�쳣��", "��ʾ", MessageBox.WARNING);
			

		} else {

			temp_revision_id = tcItemrev.getProperty("item_revision_id") + getSequenceCode(tcItemrev.getProperty("nc8_material_code"), tcItemrev.getProperty("item_revision_id"));

			if (!temp_nc8_order_number.equals("") || temp_nc8_order_number != null || !nc8_order_line_number.equals("") || nc8_order_line_number != null) {
//				saveToTCFileName = "EBOM" + "_" + nc8_order_number + "_" + nc8_drawing_no + "_" + temp_revision_id;
				saveToTCFileName = "EBOM" +  "_" + nc8_drawing_no + "_"  +nc8_material_code+"_"+ temp_revision_id;
			}else {
				saveToTCFileName = "EBOM" + "_" + nc8_drawing_no + "_"  +nc8_material_code+"_"+ temp_revision_id;			
			}

		}

		// ����汾�� �洢����ʱ��İ汾��
		System.out.println("����汾�� =" + temp_revision_id);
		ngc_utils.DoExcel.FillCell(sheet1, "O1", temp_revision_id);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G1", temp_revision_id);
		System.out.println("���ɵ�saveToTCFileName��------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F1", temp_revision_id);
		System.out.println("���ɵ�saveToTCFileName��------------- " + saveToTCFileName);

		// ����ҳ��������������ÿҳ�ĵ�Ԫ��������
		int pageBum = 1;
		if (bomLineList != null && bomLineList.size() > 0) {
			pageBum = bomLineList.size() / 34;
			if (pageBum <= 0) {
				pageBum = 1;
			}
		}

		System.out.println("����ҳ�� =" + temp_revision_id);
		// sheet1
		ngc_utils.DoExcel.FillCell(sheet1, "O2", userStr);
		System.out.println("����ҳ����------------- " + saveToTCFileName);
		// sheet2
		ngc_utils.DoExcel.FillCell(sheet2, "G2", pageBum + "");
		System.out.println("sheet2����ҳ����------------- " + saveToTCFileName);
		// sheet3
		ngc_utils.DoExcel.FillCell(sheet3, "F2", pageBum + "");
		System.out.println("sheet3����ҳ����------------- " + saveToTCFileName);
		
		saveToTCFileName = saveToTCFileName.replace("/", "-");
		saveToTCFileName = saveToTCFileName.replace("\\", "-");

		System.out.println("saveToTCFileName = " + saveToTCFileName);
		
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

		// ���ļ�д��TC������Ӧ������ĵ��汾����
		String object_name = tcItemrev.getProperty("object_name");
		String nc8_drawing_no1 = tcItemrev.getProperty("nc8_drawing_no");
		String item_revision_id = tcItemrev.getProperty("item_revision_id");
		String item_id = tcItemrev.getProperty("item_id");
		//	��item_id�Ȳ�ѯ���ݿ��Ƿ��Ѵ��ڸò�Ʒ��װ������ĵ�����ID
		String desdocId = getDesDocId(item_id);
		if (!"".equals(desdocId)) {		//�鵽������ĵ�����汾ID
			//����������汾...��ѯϵͳ�Ƿ���ڸ�����ĵ�����
			TCComponent[] componentzj = Common.CommonFinder("������汾...", "ItemID", desdocId);
			if (null != componentzj) {	//ϵͳ���ڸ�����ĵ�����
				
				TCComponentItemRevision newest = ((TCComponentItemRevision)componentzj[0]).getItem().getLatestItemRevision();
				String newestRevID = newest.getProperty("item_revision_id");
				String newestNc8_document_num2 = newest.getProperty("nc8_document_num2");
				if (newestRevID.equals(item_revision_id)) {		//����ĵ��汾���Ʒ�汾һ�£�ֱ���ù�����
 					//	�鿴����ĵ��汾�����Ƿ�����ѡ�Ĳ�Ʒ��BOMα�ļ�����
					TCComponentItemRevision tccir = getItemRevision("����ĵ��汾", desdocId, tcItemrev, "NC8_BOM");
					if(null != tccir){	//��ѡ��Ʒ��BOMα�ļ��������������ĵ��汾	
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
					}else {		//��ѡ��Ʒ��BOMα�ļ��������������ĵ��汾,��������ϵͳ��,ֱ���ù�����
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						newest.add("IMAN_specification", datasetComponent);
						tcItemrev.add("NC8_BOM", newest);
					}
					
				}else {		//����ĵ��汾���Ʒ�汾��һ�£�ԤʾҪ����
					String name = newest.getProperty("object_name");
					String description = newest.getProperty("object_desc");
					newest = newest.saveAs(item_revision_id, name, description, false, null);//�������
					newest.setProperty("nc8_document_num2",newestNc8_document_num2);
					TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
					newest.add("IMAN_specification", datasetComponent);
					tcItemrev.add("NC8_BOM", newest);
					
				}
				
			}else {		//ϵͳ�����ڸ�����ĵ�����,˵�����û���ɾ���ˣ���ô����һ������ĵ��汾������IDΪdesdocId
				TCComponentItem item = null;
				TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
				item = itemType.create(desdocId, item_revision_id, "NC8_design_doc", object_name+"_"+nc8_drawing_no1+"_"+nc8_material_code+"_��ϸ��" , "", null);
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
				TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
				itemRevision.add("IMAN_specification", datasetComponent);
				tcItemrev.add("NC8_BOM", itemRevision);
				
			}
			
			
			
		}else {	//û�в鵽����������ĵ�����汾ID
			
			//	����Des doc rev...��ѯ����ѯ����ĵ��汾
			TCComponent[] componentzj = Common.CommonFinder("Des doc rev...", "nc8_business_unit,nc8_small_class,nc8_subclass,nc8_material_code", "IBD,EBOM,EBOM" + "," + nc8_material_code);
			if (null != componentzj) {	//�鵽ϵͳ������ĵ��汾����
				// �ж����кŻ�ȡ���°汾
				TCComponentItemRevision newest = (TCComponentItemRevision)componentzj[0];
				String sequence_id = "1";
				for (int i = 0; i < componentzj.length; i++) {
					String sequence_idTemp = componentzj[i].getProperty("sequence_id");
					if (Integer.parseInt(sequence_idTemp) > Integer.parseInt(sequence_id)) {
						sequence_id = sequence_idTemp;
						newest = (TCComponentItemRevision)componentzj[i];
					}
				}
				
 				// ���Ȳ�ѯBOMα�ļ�������û������ĵ�
				TCComponentItemRevision tccir = getItemRevision("����ĵ��汾", null, tcItemrev, "NC8_BOM");
				if (null != tccir) {	//BOMα�ļ�������������ĵ��汾
					String docNc8_material_code = tccir.getProperty("nc8_material_code");
					String docRev = tccir.getProperty("item_revision_id");
					String tccirNc8_document_num2 = tccir.getProperty("nc8_document_num2");
					if (docNc8_material_code.equals(nc8_material_code) && docRev.equals(item_revision_id)) {	//ֱ�����������ĵ��汾����
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
						//	���ĵ������ID���ƷID��������
						String desDocID = tccir.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}else if (docNc8_material_code.equals(nc8_material_code)) {		//���ϱ���һ�£����汾��һ�£�˵��Ҫ����
						String name = tccir.getProperty("object_name");
						String description = tccir.getProperty("object_desc");
						tccir.saveAs(item_revision_id, name, description, false, null);//�������
						tccir.setProperty("nc8_document_num2",tccirNc8_document_num2);
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						tccir.add("IMAN_specification", datasetComponent);
//						���ĵ������ID���ƷID��������
						String desDocID = tccir.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}else {
						MessageBox.post("�ò�Ʒ����ĵ��汾��������ϱ����������Ʒ���ϱ������Բ�һ�£����޸�����ĵ��汾��������ϱ�������","����",MessageBox.ERROR);
						return;
					}
				}else {		//BOMα�ļ�������û������ĵ��汾������ϵͳ��
					String newestRevID = newest.getProperty("item_revision_id");
					String newestNc8_document_num2 = newest.getProperty("nc8_document_num2");
					if (newestRevID.equals(item_revision_id)) {		//���Ʒ�汾һ�£���ֱ����
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						newest.add("IMAN_specification", datasetComponent);
						tcItemrev.add("NC8_BOM", newest);
						// ���ĵ������ID���ƷID��������
						String desDocID = newest.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}else {		//���Ʒ�汾��һ�£�������
						String name = newest.getProperty("object_name");
						String description = newest.getProperty("object_desc");
						newest = newest.saveAs(item_revision_id, name, description, false, null);//�������
						newest.setProperty("nc8_document_num2",newestNc8_document_num2);
						TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
						newest.add("IMAN_specification", datasetComponent);
						tcItemrev.add("NC8_BOM", newest);
//						���ĵ������ID���ƷID��������
						String desDocID = newest.getProperty("item_id");
						String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
						insertDesDocId(item_id, desDocID, insertUser);
					}
				}
				
			}else {		//�鵽ϵͳû������ĵ��汾����
				TCComponentItemType itemType = (TCComponentItemType) session.getTypeComponent("Item");
				String newID = itemType.getNewID();
				TCComponentItem item = itemType.create(newID, item_revision_id, "NC8_design_doc", object_name+"_"+nc8_drawing_no1+"_"+nc8_material_code+"_��ϸ��" , "", null);
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
				TCComponentDataset datasetComponent = ReportCommon.setDatasetFileToTC(OutFileName, "MS Excel", "excel", saveToTCFileName);
//				itemRevision.add("IMAN_reference", datasetComponent);
				itemRevision.add("IMAN_specification", datasetComponent);
				tcItemrev.add("NC8_BOM", itemRevision);
				
				//	������ĵ��汾�����ID���ڲ�Ʒ��
				String insertUser = ((TCComponentPerson) (session.getUser().getReferenceProperty("person"))).toString();
				insertDesDocId(item_id, newID, insertUser);
				
			}

			
		}
		
//		TCComponentDataset datasetComponent = ReportCommon.hasDataset(datasetType, datasetName, relationObject, relationType);
//		ReportCommon.createOrUpdateExcel(OutFileName, saveToTCFileName, itemRevision, itemRevision.getType(), true);
		System.out.println("д��TC�ɹ�--------------");

		// ��Ԥ��
		Runtime.getRuntime().exec("cmd /c start " + OutFileName);
		MessageBox.post("����������ϣ�", "��ʾ", 2);

	}

	/**
	 * �������ݼ�
	 * 
	 * @param componentDataset
	 *            ���ݼ�����
	 * @param namedRefName
	 *            ���ݼ�����
	 * @param localDir
	 *            ����Ŀ¼
	 * @return
	 */
	public synchronized static String[] FileToLocalDir(TCComponentDataset componentDataset, String namedRefName, String localDir) {
		try {
			// ��ȡ����·��
			File dirObject = new File(localDir);
			if (!dirObject.exists()) {
				dirObject.mkdirs();
			}

			componentDataset = componentDataset.latest();

			// ע�⣺��������[������]��ͬ���ļ����ܴ��ڶ��
			String namedRefFileName[] = componentDataset.getFileNames(namedRefName);
			if ((namedRefFileName == null) || (namedRefFileName.length == 0)) {
				 Common.ShowTcErrAndMsg("���ݼ�<" + componentDataset.toString() +
				 ">û�ж�Ӧ����������!");
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
			 Common.ShowTcErrAndMsg("���ݼ�<" + componentDataset.toString() +
			 ">���ô���!");
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
				// if (BOMLine.getProperty("bl_uom").equals("ÿ��")) {
				// if (!BOMLine.getProperty("bl_quantity").equals("")) {
				// if (Integer.valueOf(BOMLine.getProperty("bl_quantity"))
				// .intValue() > 1) {
				// BOMLine.unpack();
				// }
				// }
				// }
				// �ж��Ƿ񷢲�
				/*
				 * boolean idHasStatus =
				 * TcUtils.idHasStatus(BOMLine.getItemRevision());
				 * System.out.println("���Ƿ񷢲���idHasStatus = " + idHasStatus);
				 * if(idHasStatus != false){ ColletcBOMView(BOMLine, Level + 1,
				 * bomLineListTest); }
				 */
				if (!BOMLine.isRoot()) {
					TCComponentItemRevision bRevision = BOMLine.getItemRevision();
					String bString = bRevision.getProperty("object_name");
					// �ж��Ƿ�չ��
					String NC8_autoExpand_true = BOMLine.getProperty("NC8_autoExpand_true");
					System.out.println("��" + bString + "���Ƕ��㣬�Ƿ�չ��Ϊ=" + NC8_autoExpand_true + "��");
					if (!"��".equals(NC8_autoExpand_true)) {
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
	
	
	private String getSequenceCode(String sequenceName, String revision) {

		if (sequenceName != null && sequenceName.length() != 0) {

			System.out.println("�ϴ���sequenceNameΪ------------- " + sequenceName);
			String sequenceCode = String.valueOf(JDBCUtils.querySequenceCode(sequenceName, revision));
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
	
	

	private void fillValue(TCComponentBOMLine bomLine, TCComponentItemRevision bomLineRevision, Sheet sheet1, Integer level, Sheet sheet2)
			throws TCException {
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
		String nc8_material_code_parent = "";
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

		String object_type = bomLineRevision.getTCProperty("object_type").getStringValue();
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
				String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
				String nc8_specification = bomLineRevision.getProperty("nc8_specification");
				System.out.println("�����--------------" + nc8_specification);
				daihao = nc8_drawing_no + " " + nc8_specification;
				System.out.println("�����š�--------------" + daihao);
				// �㼶
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// Ӣ������
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					}

				} else {
					System.out.println("��û���ϲ��������");
					nc8_material_code_parent = "";
				}

				// ���� (����δд��)

				/**
				 * nc8_order_number =
				 * bomLineRevision.getProperty("nc8_order_number");
				 * System.out.println("�����ϡ�--------------" + nc8_drawing_no);
				 */

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("��������--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// ����
				nc8_weight = bomLineRevision.getProperty("nc8_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);

				// ľģͼ�� item ���ԣ��о��ã�
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			} else if ("NC8_CastingRevision".equals(object_type) || "NC8_ForgingsRevision".equals(object_type)
					|| "NC8_WeldingRevision".equals(object_type) || "NC8_SectionRevision".equals(object_type)) {
				/**
				 * ԭ����
				 */
				System.out.println("���ö���Ϊԭ���ϡ�");
				// ���� nc8_Standard+�� ��+nc8_specification
				String nc8_Standard = bomLineRevision.getProperty("nc8_Standard");
				System.out.println("����׼��--------------" + nc8_Standard);
				String nc8_Specification = bomLineRevision.getProperty("nc8_Specification");
				System.out.println("�����--------------" + nc8_Specification);
				String drawing_no3 = bomLineRevision.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_Specification;
				daihao = drawing_no3;//2019/01/10���ģ���nc8_Standard��nc8_Specificationƴ�ӵ�ֵ��Ϊnc8_drawing_no3��ֵ
				System.out.println("�����š�--------------" + daihao);
				// �㼶
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// Ӣ������
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = bomLineRevision.getProperty("nc8_Materialnumber");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("��û���ϲ��������");
					nc8_material_code_parent = "";
				}

				// ����
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("��������--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// ����
				nc8_weight = bomLineRevision.getProperty("nc8_net_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			} else if ("NC8_AssistantMatRevision".equals(object_type)) {
				/**
				 * ����
				 */
				System.out.println("���ö���Ϊ���ϡ�");
				// ���� nc8_Standard +�� ��+ nc8_model+�� ��+ nc8_Specification
				String nc8_Standard = bomLineRevision.getProperty("nc8_Standard");
				System.out.println("����׼��--------------" + nc8_Standard);
				String nc8_model = bomLineRevision.getProperty("nc8_model");
				System.out.println("���ͺš�--------------" + nc8_model);
				String nc8_specification = bomLineRevision.getProperty("nc8_Specification");
				System.out.println("�����--------------" + nc8_specification);
				String drawing_no3 = bomLineRevision.getProperty("nc8_drawing_no3");
				//daihao = nc8_Standard + " " + nc8_model + " " + nc8_specification;
				daihao = drawing_no3;//2019/01/10���ģ���nc8_Standard��nc8_Specificationƴ�ӵ�ֵ��Ϊnc8_drawing_no3��ֵ
				System.out.println("�����š�--------------" + daihao);
				// �㼶
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// Ӣ������
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = bomLineRevision.getProperty("nc8_Materialnumber");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("��û���ϲ��������");
					nc8_material_code_parent = "";
				}

				// ����
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
				/**
				 * ���Ȼ�ȡ������������������Ϊ�յ�ʱ����ȥ��ȡ����
				 */
				String nc8_assist_number = bomLine.getProperty("NC8_Assist_number");
				System.out.println("���������� = " + nc8_assist_number + "��");
				if ("".equals(nc8_assist_number) || nc8_assist_number == null) {
					String bl_quantity_bak = bomLine.getProperty("bl_quantity");
					System.out.println("��������--------------" + bl_quantity_bak);
					if (bl_quantity_bak == null || bl_quantity_bak.length() == 0) {
						bl_quantity = "1";
					}else {
						bl_quantity = bl_quantity_bak;
					}
				}else {
					bl_quantity = nc8_assist_number;
				}
				System.out.println("��excel������--------------" + bl_quantity);
				

				// ����
				nc8_weight = bomLineRevision.getProperty("nc8_net_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			} else if ("NC8_test_piecesRevision".equals(object_type)) {
				/**
				 * �����
				 */
				System.out.println("���ö���Ϊ�������");
				// ���� nc8_drawing_no
				String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("�����š�--------------" + daihao);
				// �㼶
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// Ӣ������
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("��û���ϲ��������");
					nc8_material_code_parent = "";
				}

				// ����
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("��������--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}

				// ����
				nc8_weight = bomLineRevision.getProperty("nc8_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			} else if ("NC8_purchasedRevision".equals(object_type)) {
				/**
				 * �⹺��
				 */
				System.out.println("���ö���Ϊ�⹺����");
				// ���� nc8_drawing_no
				String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
				daihao = nc8_drawing_no;
				System.out.println("�����š�--------------" + daihao);
				// �㼶
				bl_sequence_no = bomLine.getProperty("bl_sequence_no");
				System.out.println("���㼶��" + bl_sequence_no);
				// �汾
				item_revision_id = bomLineRevision.getProperty("item_revision_id");
				System.out.println("���汾��--------------" + item_revision_id);
				// ��������
				object_name = bomLineRevision.getProperty("object_name");
				System.out.println("���������ơ�--------------" + object_name);
				// Ӣ������
				nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
				// ���ϱ���
				nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
				System.out.println("�����ϱ��롿--------------" + nc8_material_code);
				// �������
				TCComponentBOMLine parentBomLine = bomLine.parent();
				if (parentBomLine != null) {
					TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
					if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
						nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
						System.out.println("��������롿--------------" + nc8_material_code_parent);
					}
				} else {
					System.out.println("��û���ϲ��������");
					nc8_material_code_parent = "";
				}

				/**
				 * ���� nc8_material+nc8_grade+nc8_hardness_level 2018-08-21�޸�
				 * 1.��ϸ�����⹺��ITEM�ġ����ϡ���ȡITEM���ԣ�����+���ܵȼ�+Ӳ�ȵȼ���
				 * 2.��ϸ�����⹺��ITEM�ġ���ע����ȡITEM���ԣ�������+BOM��ע��
				 */
				nc8_material = bomLineRevision.getProperty("nc8_material");
				System.out.println("�����ʡ�--------------" + nc8_material);
				String nc8_grade = bomLineRevision.getProperty("nc8_grade");
				System.out.println("�����ܵȼ���--------------" + nc8_grade);
				String nc8_hardness_level = bomLineRevision.getProperty("nc8_hardness_level");
				System.out.println("��Ӳ�ȵȼ���--------------" + nc8_hardness_level);
				nc8_material = nc8_material + " " + nc8_grade + " " + nc8_hardness_level;
				System.out.println("�����ϡ�--------------" + nc8_material);

				// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
				bl_quantity = bomLine.getProperty("bl_quantity");
				System.out.println("��������--------------" + bl_quantity);
				if (bl_quantity == null || bl_quantity.length() == 0) {
					bl_quantity = "1";
				}
				// ����
				nc8_weight = bomLineRevision.getProperty("nc8_weight");
				System.out.println("�����ء�--------------" + nc8_weight);
				// ��ע nc8_feature_set- nc8_grade - nc8_hardness_level
				// +NC8_BOM_remark
				String nc8_feature_set = bomLineRevision.getProperty("nc8_feature_set");
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
				NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
				System.out.println("��BOM��ע��--------------" + NC8_BOM_remark);
				NC8_BOM_remark = (nc8_feature_set + " " + NC8_BOM_remark).trim();
				System.out.println("����ע��--------------" + NC8_BOM_remark);
				// ľģͼ��
				if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
					nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
					System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
				}

				// ��������

			}

		} else {
			System.out.println("���ö���Ϊ��ͨ����");
			// ���� nc8_drawing_no
			String nc8_drawing_no = bomLineRevision.getProperty("nc8_drawing_no");
			System.out.println("��ͼ�š�--------------" + nc8_drawing_no);
			daihao = nc8_drawing_no;
			System.out.println("�����š�--------------" + daihao);
			// �㼶
			bl_sequence_no = bomLine.getProperty("bl_sequence_no");
			System.out.println("���㼶��" + bl_sequence_no);
			// �汾
			item_revision_id = bomLineRevision.getProperty("item_revision_id");
			System.out.println("���汾��--------------" + item_revision_id);
			// ��������
			object_name = bomLineRevision.getProperty("object_name");
			System.out.println("���������ơ�--------------" + object_name);
			// Ӣ������
			nc8_part_name = bomLineRevision.getProperty("nc8_part_name");
			System.out.println("��Ӣ�����ơ�--------------" + nc8_part_name);
			// ���ϱ���
			nc8_material_code = bomLineRevision.getProperty("nc8_material_code");
			System.out.println("�����ϱ��롿--------------" + nc8_material_code);
			// �������
			TCComponentBOMLine parentBomLine = bomLine.parent();
			if (parentBomLine != null) {
				TCComponentItemRevision parentItemRevision = parentBomLine.getItemRevision();
				if (parentItemRevision.isValidPropertyName("nc8_material_code")) {
					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_code");
					System.out.println("��������롿--------------" + nc8_material_code_parent);
				} else if (parentItemRevision.isValidPropertyName("nc8_material_number")) {
					nc8_material_code_parent = parentItemRevision.getProperty("nc8_material_number");
					System.out.println("��������롿--------------" + nc8_material_code_parent);
				}
			} else {
				System.out.println("��û���ϲ��������");
				nc8_material_code_parent = "";
			}

			// ���� nc8_material+nc8_grade+nc8_hardness_level
			nc8_material = bomLineRevision.getProperty("nc8_material");
			System.out.println("�����ϡ�--------------" + nc8_material);
			// ���� ����bomline���ԣ�ֵΪ0ʱ����ʾ1��
			bl_quantity = bomLine.getProperty("bl_quantity");
			System.out.println("��������--------------" + bl_quantity);
			if (bl_quantity == null || bl_quantity.length() == 0) {
				bl_quantity = "1";
			}
			// ����
			nc8_weight = bomLineRevision.getProperty("nc8_weight");
			System.out.println("�����ء�--------------" + nc8_weight);
			// ��ע
			NC8_BOM_remark = bomLine.getProperty("NC8_BOM_remark");
			System.out.println("����ע��--------------" + NC8_BOM_remark);
			// ľģͼ��
			if (bomLineRevision.isValidPropertyName("nc8_wood_pattern") && (bomLineRevision.getProperty("nc8_wood_pattern").length() != 0)) {
				nc8_wood_pattern = bomLineRevision.getProperty("nc8_wood_pattern");
				System.out.println("��ľģͼ�š�--------------" + nc8_wood_pattern);
			}

			// ��������

		}
		// ��ţ���ˮ��������
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
		// System.out.println("�����ղ㼶��" + bl_sequence_no);
		rowNum++;
		number++;

		TCComponent[] relatedComponents = bomLineRevision.getRelatedComponents("IMAN_specification");
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
		
		String nc8_firstused_products = bomLineRevision.getProperty("nc8_firstused_products");
		if (nc8_firstused_products.equals(whole_nc8_drawing_no)) {
			boolean isRoot = bomLine.isRoot();
			
			if (isCreate || isRoot) {
				/**
				 * ��ƷͼֽĿ¼
				 */
				// ͼ��
				String nc8_drawing_no_product = bomLineRevision.getProperty("nc8_drawing_no");
				System.out.println("����ƷͼֽĿ¼-ͼ�š�--------------" + nc8_drawing_no_product);
				ngc_utils.DoExcel.FillCell(sheet2, "A" + productRowNum, nc8_drawing_no_product);

				// �汾
				String item_revision_id_product = bomLineRevision.getProperty("item_revision_id");
				System.out.println("����ƷͼֽĿ¼-�汾��--------------" + item_revision_id_product);
				ngc_utils.DoExcel.FillCell(sheet2, "B" + productRowNum, item_revision_id_product);

				// ��������
				String object_name_product = bomLineRevision.getProperty("object_name");
				System.out.println("����ƷͼֽĿ¼-�������ơ�--------------" + object_name_product);
				ngc_utils.DoExcel.FillCell(sheet2, "C" + productRowNum, object_name_product);

				// Ӣ������
				String nc8_part_name_product = bomLineRevision.getProperty("nc8_part_name");
				System.out.println("����ƷͼֽĿ¼-Ӣ�����ơ�--------------" + nc8_part_name_product);
				ngc_utils.DoExcel.FillCell(sheet2, "D" + productRowNum, nc8_part_name_product);

				// ͼ��
				String nc8_drawing_size_product = bomLineRevision.getProperty("nc8_drawing_size");
				System.out.println("����ƷͼֽĿ¼-ͼ����--------------" + nc8_drawing_size_product);
				ngc_utils.DoExcel.FillCell(sheet2, "E" + productRowNum, nc8_drawing_size_product);

				// ҳ��
				String nc8_pages_product = bomLineRevision.getProperty("nc8_pages");
				System.out.println("����ƷͼֽĿ¼-ҳ����--------------" + nc8_pages_product);
				ngc_utils.DoExcel.FillCell(sheet2, "F" + productRowNum, nc8_pages_product);

				// ��ע
				String nc8_remarks_product = bomLineRevision.getProperty("nc8_remarks");
				System.out.println("����ƷͼֽĿ¼-��ע��--------------" + nc8_remarks_product);
				ngc_utils.DoExcel.FillCell(sheet2, "G" + productRowNum, nc8_remarks_product);

				productRowNum++;
			}
		}
		
					
	}
	
	
	
	 public static boolean idHasStatus(TCComponent component) throws TCException{
			boolean flag =false;
		TCComponent[] components = 	component.getReferenceListProperty("release_status_list");
		System.out.println("״̬2"+components.length);
		for (int i = 0; i < components.length; i++) {
			String type = components[i].getProperty("object_name");
			System.out.println("״̬����"+type);
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
			// �ж��Ƿ�չ��
			String NC8_autoExpand_true = childBOMLine.getProperty("NC8_autoExpand_true");
			//if ("��".equals(NC8_autoExpand_true)) {
				//�жϡ���չ���������е�ֵΪ���ǡ�
				String NC8_Y_or_N_Expand = childBOMLine.getProperty("NC8_Y_or_N_Expand");
				if ("��".equals(NC8_Y_or_N_Expand)) {
					ColletcBOMView(childBOMLine, Level + 1, bomLineListTest);
				}
			//}
			checkYOrNExpand(childBOMLine, Level + 1, bomLineListTest);
		}
		
	}
	
	
	
	//�ж϶������Ƿ�����ض���ϵ��������汾
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
		
		
//		String bigClass=tcItemrev.getProperty("nc8_big_class");
		System.out.println("bigClassΪ--------------" + bigClass);
	   	
		
		//С��
		TCProperty smallClass = tcItemrev.getTCProperty("nc8_small_class");
//		String smallClass="VB";
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
		
//		try {
			//tcItemrev.setProperty("nc8_document_num2", docuNum);//����ʵ��Ϊnc8_document_num��ֻ��״̬��������nc8_document_num2Ϊ��֮��������ȵ����ԣ�BMIDE������Ϊ���ؿ�д
			//MessageBox.post("���ɵ��ĵ����Ϊ��"+docuNum,"��ʾ",MessageBox.INFORMATION);
			return docuNum;
//		} catch (Exception e1) {
//			MessageBox.post(e1.toString(),"��ʾ",MessageBox.WARNING);
//			e1.printStackTrace();
//			return "";
//		}
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

	// ���û���������ĵ��汾ʱ�������ݿ�����ƷID������ĵ�ID�Ķ�Ӧ��ϵ
	private void insertDesDocId(String productId, String desdocId, String insertUser) {
		

		Connection conn = null;// ����һ�����ݿ�����
		PreparedStatement pre = null; // ����Ԥ����������
		ResultSet result = null;
		try {

			Class.forName("oracle.jdbc.driver.OracleDriver");// ����Oracle��������
			System.out.println("��ʼ�����������ݿ⣡");
			ExportCommon ec = new ExportCommon(); // ���ڻ�ȡ��Ҫ���ӵ����ݿ����Ϣ���������ݿ��ַ���û���������
			String url = ec.getOracle_url_dev();// ���ݿ��ַ
			String user = ec.getOracle_user();// �û���
			String password = ec.getOracle_password();// ����
			conn = DriverManager.getConnection(url, user, password);// ��ȡ����
			System.out.println("���ӳɹ���");

			// sql���
			String sql = "insert into PLM_PRODUCT_DESDOC_RELATION (PRODUCT_ID, DESDOC_ID, INSERT_USER)" + " values(?, ?, ?)";
			pre = conn.prepareStatement(sql);// ʵ����Ԥ�������
			pre.setString(1, productId);// ���ò�����ǰ���1��ʾ�����������������Ǳ�������������
			pre.setString(2, desdocId);// ���ò�����ǰ���2��ʾ�����������������Ǳ�������������
			pre.setString(3, insertUser);// ���ò�����ǰ���3��ʾ�����������������Ǳ�������������

			result = pre.executeQuery();// ִ�в�ѯ��ע�������в���Ҫ�ټӲ���

		}
		catch (Exception e1) {
			e1.printStackTrace();
		}
		finally {
			try {
				// ��һ������ļ�������رգ���Ϊ���رյĻ���Ӱ�����ܡ�����ռ����Դ
				// ע��رյ�˳�����ʹ�õ����ȹر�
				if (result != null)
					result.close();
				if (pre != null)
					pre.close();
				if (conn != null)
					conn.close();
				System.out.println("���ݿ������ѹرգ�");
			}
			catch (Exception e2) {
				e2.printStackTrace();
			}
		}
	}
	
	/*
	// ���²�ƷID������ĵ�ID�Ķ�Ӧ��ϵ
	private void updateDesDocId(String productId, String desdocId, String updateUser) {
		

		Connection conn = null;// ����һ�����ݿ�����
		PreparedStatement pre = null; // ����Ԥ����������
		ResultSet result = null;
		try {

			Class.forName("oracle.jdbc.driver.OracleDriver");// ����Oracle��������
			System.out.println("��ʼ�����������ݿ⣡");
			ExportCommon ec = new ExportCommon(); // ���ڻ�ȡ��Ҫ���ӵ����ݿ����Ϣ���������ݿ��ַ���û���������
			String url = ec.getOracle_url_dev();// ���ݿ��ַ
			String user = ec.getOracle_user();// �û���
			String password = ec.getOracle_password();// ����
			conn = DriverManager.getConnection(url, user, password);// ��ȡ����
			System.out.println("���ӳɹ���");

			// sql���
			String sql = "update PLM_PRODUCT_DESDOC_RELATION set DESDOC_ID = ?, UPDATE_USER = ? where PRODUCT_ID = ?";
			pre = conn.prepareStatement(sql);// ʵ����Ԥ�������
			pre.setString(1, desdocId);// ���ò�����ǰ���1��ʾ�����������������Ǳ�������������
			pre.setString(2, updateUser);// ���ò�����ǰ���2��ʾ�����������������Ǳ�������������
			pre.setString(3, productId);// ���ò�����ǰ���3��ʾ�����������������Ǳ�������������

			result = pre.executeQuery();// ִ�в�ѯ��ע�������в���Ҫ�ټӲ���

		}
		catch (Exception e1) {
			e1.printStackTrace();
		}
		finally {
			try {
				// ��һ������ļ�������رգ���Ϊ���رյĻ���Ӱ�����ܡ�����ռ����Դ
				// ע��رյ�˳�����ʹ�õ����ȹر�
				if (result != null)
					result.close();
				if (pre != null)
					pre.close();
				if (conn != null)
					conn.close();
				System.out.println("���ݿ������ѹرգ�");
			}
			catch (Exception e2) {
				e2.printStackTrace();
			}
		}
	}
*/	
	// ��ѯ���ݿ��Ƿ����в�ƷID����Ӧ������ĵ�ID
	private String getDesDocId(String productId) {
		
		String resultId = "";

		Connection conn = null;// ����һ�����ݿ�����
		PreparedStatement pre = null; // ����Ԥ����������
		ResultSet result = null;
		try {

			Class.forName("oracle.jdbc.driver.OracleDriver");// ����Oracle��������
			System.out.println("��ʼ�����������ݿ⣡");
			ExportCommon ec = new ExportCommon(); // ���ڻ�ȡ��Ҫ���ӵ����ݿ����Ϣ���������ݿ��ַ���û���������
			String url = ec.getOracle_url_dev();// ���ݿ��ַ
			String user = ec.getOracle_user();// �û���
			String password = ec.getOracle_password();// ����
			conn = DriverManager.getConnection(url, user, password);// ��ȡ����
			System.out.println("���ӳɹ���");

			// sql���
			String sql = "select DESDOC_ID from PLM_PRODUCT_DESDOC_RELATION where PRODUCT_ID = ?";
			pre = conn.prepareStatement(sql);// ʵ����Ԥ�������
			pre.setString(1, productId);// ���ò�����ǰ���1��ʾ�����������������Ǳ�������������

			result = pre.executeQuery();// ִ�в�ѯ��ע�������в���Ҫ�ټӲ���
			// ����鵽�ˣ�����ͻ᷵��true������᷵��false
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
				// ��һ������ļ�������رգ���Ϊ���رյĻ���Ӱ�����ܡ�����ռ����Դ
				// ע��رյ�˳�����ʹ�õ����ȹر�
				if (result != null)
					result.close();
				if (pre != null)
					pre.close();
				if (conn != null)
					conn.close();
				System.out.println("���ݿ������ѹرգ�");
			}
			catch (Exception e2) {
				e2.printStackTrace();
			}
		}
		return resultId;
	}
	
	
	
	//����ĵ�����ʱд������ĵ����,���Ƶ�com.uds.drawingNumber.manually.handlers.CopyItemRevProperty�Ĵ���
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
				// �����������ж�������� : p_value_code
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
			
			// ���ѡ���item�������ĵ�
			// �ж��ĵ����� : p_object_type
			// NC8_design_doc ����ĵ�
			// NC8_general_doc ͨ���ĵ�
			// NC8_process_doc �����ĵ�
			// NC8_standard_doc ��׼�ĵ�
			if (p_object_type.equals("NC8_design_doc")
					|| p_object_type.equals("NC8_general_doc")
					|| p_object_type.equals("NC8_process_doc")
					|| p_object_type.equals("NC8_standard_doc")) {
				// initUI();
				tcItemrev.setProperty("nc8_document_num2",p_document_no);
				MessageBox.post("�ɹ�д�� �ĵ���� : " + p_document_no, "���Ƴɹ�",MessageBox.INFORMATION);
			}
			
		} catch (TCException e) {
			// TODO Auto-generated catch block
			MessageBox.post(e.getDetailsMessage() + "��ȡ��ǰ�汾���ĵ����ʧ��, ����ϵ����Ա��ѯ���౨������! ", "����",MessageBox.ERROR);
			e.printStackTrace();
		}
		return false;
	}
	
	

}
