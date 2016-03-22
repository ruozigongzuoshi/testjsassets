import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class excel4hb {
	
	private static String Path = "C:\\Users\\jishu\\Desktop\\��ʳ���-ÿ������-3��-����Դ.xls";
	private static String NewPath = "C:\\Users\\jishu\\Desktop\\��ʳ���-ÿ������-3��-����Դ2.xls";
	private static String ConfigPath = "C:\\Users\\jishu\\Desktop\\������������.xls";
	
	

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		jxl.Workbook readwb = null;
		try {
			InputStream instream = new FileInputStream(Path);
			readwb = Workbook.getWorkbook(instream);
			int countSheet = readwb.getNumberOfSheets();
			int readSheet = countSheet - 1;
			// Sheet���±��Ǵ�0��ʼ
			// ��ȡ��һ��Sheet��
			String[][][] table = inputMemory(readwb,readSheet);
			String[][] configname = getConfigData(ConfigPath);
			String[][] writetable = new String[readwb.getSheet(0).getRows()][30];
			System.out.println("**������" +  countSheet);
			System.out.println("##�ؼ��ʷ���");
			for (int k = 0; k < readSheet; k++) {
				Sheet readsheet = readwb.getSheet(k);
				int rsRows = readsheet.getRows();
				System.out.println("**������" + (k + 1) + "�ؼ���");
				
				
				switch(k){
				case 0:
				{
					for (int i = 3; i < rsRows; i++) {
						
						writetable[i][0] = getDate(table[k][i][0]);//����
						writetable[i][1] = changeName(configname,table[k][i][1]);//��������
						writetable[i][2] = table[k][i][2];//�豸����
						writetable[i][3] = getdatafromtable(
								table, table[k][i][0], 0, table[k][i][1], 17, 21);//�����û���
						writetable[i][4] = getdatafromtable(
								table, table[k][i][0], 0, table[k][i][1], 17, 18);//���ո���
						writetable[i][5] = getdatafromtable(
								table, table[k][i][0], 0, table[k][i][1], 7, 10);//��������
						writetable[i][6] = getdatafromtable(
								table, table[k][i][0], 0, table[k][i][1], 7, 12);//��������
						writetable[i][7] = getdatafromtable(
								table, table[k][i][0], 0, table[k][i][1], 7, 14);//��������
						
						
//					System.out.println("���ڣ�" + writetable[i][0] + "������"+ writetable[i][1]
//							+ "���"+ writetable[i][2] + "����������"+ writetable[i][3]
//									+ "�����ܶ"+ writetable[i][4] + "�������棺"+ writetable[i][5]
//											+ "�������棺"+ writetable[i][6] + "�������棺"+ writetable[i][7]);
					}
					break;
				}
				case 1:
				{
					for(int i = 1; i < rsRows; i++){
						String d_time = getDate(table[k][i][0]);//����
						String q_name = changeName(configname,table[k][i][6]);//��������
						
						int curRow = getCurRowfromWriteTable(writetable,d_time,q_name);	
						if(curRow == 0){
							System.out.println("��" + (k+1) + "�ű�\t����:" + table[k][i][0] + "\t����:" + table[k][i][6] + "û����dataeye�ҵ���");
						}
						
						
						if(writetable[curRow][8] == null){
							writetable[curRow][8] = "0";//��ʼ�����ѽ��
						}
						
						if(writetable[curRow][9] == null){
							writetable[curRow][9] = "0";//��ʼ�������û���
						}

						if(writetable[curRow][10] == null){
							writetable[curRow][10] = "0";//��ʼ�����Ѵ���
						}
						if(writetable[curRow][11] == null){
							writetable[curRow][11] = "0";//��ʼ������arppu
						}
						
						writetable[curRow][8] = String.valueOf(Double.parseDouble(writetable[curRow][8]) 
									+ Double.parseDouble(table[k][i][9]));//������
						writetable[curRow][9] = String.valueOf(Double.parseDouble(writetable[curRow][9]) 
									+ Double.parseDouble(table[k][i][10]));//�����û���
						writetable[curRow][10] = String.valueOf(Double.parseDouble(writetable[curRow][10]) 
									+ Double.parseDouble(table[k][i][11]));//���Ѵ���
						writetable[curRow][11] = String.valueOf(Double.parseDouble(writetable[curRow][11]) 
									+ Double.parseDouble(table[k][i][26]));//��arppu
						
					}
					break;
				}
				case 2:
				{
					for(int i = 1; i < rsRows; i++){
						String d_time = getDate(table[k][i][0]);//����
						String q_name = changeName(configname,table[k][i][3]);//��������
						
						int curRow = getCurRowfromWriteTable(writetable,d_time,q_name);	
						if(curRow == 0){
							System.out.println("��" + (k+1) + "�ű�\t����:" + table[k][i][0] + "\t����:" + table[k][i][3] + "û����dataeye�ҵ���");
						}
						
						if(writetable[curRow][12] == null){
							writetable[curRow][12] = "0";//��ʼ��������
						}
						
						if(writetable[curRow][13] == null){
							writetable[curRow][13] = "0";//��ʼ�������û���
						}

						if(writetable[curRow][14] == null){
							writetable[curRow][14] = "0";//��ʼ�����Ѵ���
						}
						if(writetable[curRow][15] == null){
							writetable[curRow][15] = "0";//��ʼ����arppu
						}
						
						writetable[curRow][12] = String.valueOf(Double.parseDouble(writetable[curRow][12]) 
									+ Double.parseDouble(table[k][i][8]));//������
						writetable[curRow][13] = String.valueOf(Double.parseDouble(writetable[curRow][13]) 
									+ Double.parseDouble(table[k][i][11]));//�����û���
						writetable[curRow][14] = String.valueOf(Double.parseDouble(writetable[curRow][14]) 
									+ Double.parseDouble(table[k][i][12]));//���Ѵ���
						writetable[curRow][15] = String.valueOf(Double.parseDouble(writetable[curRow][15]) 
									+ Double.parseDouble(table[k][i][13]));//��arppu
						
						
					}
					break;
				}
				case 3:
				{
					for(int i = 1; i < rsRows; i++){
						String d_time = getDate(table[k][i][1]);//����
						String q_name = changeName(configname,table[k][i][2]);//��������
						
						int curRow = getCurRowfromWriteTable(writetable,d_time,q_name);
						if(curRow == 0){
							System.out.println("��" + (k+1) + "�ű�\t����:" + table[k][i][0] + "\t����:" + table[k][i][2] + "û����dataeye�ҵ���");
						}
						
						if(writetable[curRow][16] == null){
							writetable[curRow][16] = "0";//��ʼ����Ϸ����
						}
						
						if(writetable[curRow][17] == null){
							writetable[curRow][17] = "0";//��ʼ�������û���
						}

						if(writetable[curRow][18] == null){
							writetable[curRow][18] = "0";//��ʼ��arppu
						}
						
						
						writetable[curRow][16] = String.valueOf(Double.parseDouble(writetable[curRow][16]) 
									+ Double.parseDouble(table[k][i][6]));//��Ϸ����
						writetable[curRow][17] = String.valueOf(Double.parseDouble(writetable[curRow][17]) 
									+ Double.parseDouble(table[k][i][10]));//�����û���
						writetable[curRow][18] = String.valueOf(Double.parseDouble(writetable[curRow][18]) 
									+ Double.parseDouble(table[k][i][13]));//arppu
						
						
						
					}
					break;
				}
				default:break;
				}
				
			}
			outXls(readwb, table, writetable, readSheet);
			
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			readwb.close();
		}
	
	}
	
	public static String getDate(String ss){
		DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
		DateFormat df2 = new SimpleDateFormat("yyyy/MM/dd");
		Date date = null;
		
		try {
			date = df.parse(ss);
		} catch (Exception e) {
			try{
				date = df2.parse(ss);
			}catch(Exception e2){
				
				if(Long.valueOf(ss).longValue() > 99999){
				return ss.substring(2,4).trim() + "-" + ss.substring(4,6).trim() + "-"
					+ ss.substring(6,8).trim();
				}	
				Long l = (Long.valueOf(ss).longValue() - 25569) * 86400000;
				date = new Date(l);
				
				}
			
		}
		return df.format(date).substring(2);
	}
	
	public static String[][] getConfigData(String ss){
		jxl.Workbook readwb = null;
		String[][] config_ids = null;
		try {
			InputStream instream = new FileInputStream(ss);
			readwb = Workbook.getWorkbook(instream);
			int countSheet = readwb.getNumberOfSheets();
			// Sheet���±��Ǵ�0��ʼ
			// ��ȡ��һ��Sheet��
			Sheet readsheet = readwb.getSheet(0);
			int rsRows = readsheet.getRows();
			int rsCells = readsheet.getColumns();
			config_ids = new String[rsRows][];
			for (int k = 0; k < rsRows; k++) {
				config_ids[k] = new String[rsCells];
				for (int i = 0; i < rsCells; i++) {
					config_ids[k][i] = readsheet.getCell(i, k).getContents();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			readwb.close();
		}
		System.out.println("��ȡ�������Ƴɹ�");
		return config_ids;
	}
	
	public static String[][][] inputMemory(Workbook readwb, int readSheet) {
		// ���ص��ڴ���
		String[][][] table = new String[readSheet][][];
		for (int i = 0; i < readSheet; i++) {
			Sheet st = readwb.getSheet(i);
			int r = st.getRows();
			int c = st.getColumns();
			table[i] = new String[r][c];
			for (int j = 0; j < r; j++) {
				for (int k = 0; k < c; k++) {
					table[i][j][k] = st.getCell(k, j).getContents();
				}
			}
		}
		System.out.printf("д���ڴ�ɹ�");
		return table;
	}
	
	public static String getdatafromtable(String[][][] table, String id, int i_id, String name, int i_name, int num){
		String data = "0";
		for(int i = 0; i < table[0].length; i++){
			if(table[0][i][i_id].equalsIgnoreCase(id)&&table[0][i][i_name].equalsIgnoreCase(name)){
				data = table[0][i][num];
				break;
			}
		}
		return data;
	}
	
	public static int getCurRowfromWriteTable(String[][] writetable, String id, String name){
		int curRow = 0;
		for(int i = 3; i < writetable.length; i++){
			if(writetable[i][0].equalsIgnoreCase(id)&&(writetable[i][1].equalsIgnoreCase(name)))
			{
				curRow = i;
				break;
			}
			
		}
//		if(curRow == 0){
//			System.out.println(id + name + "����δ�ҵ�");
//		}
		return curRow;
	}
	
	public static String changeName(String[][] configname, String ss){
		for(int i = 0 ; i < configname.length; i++){
			for(int j = 0; j < configname[i].length; j++){
				if(ss.equalsIgnoreCase(configname[i][j])){
					//System.out.println(ss+ "����:" + configname[i][j]);
					return configname[i][0];
				}
			}
		}
		//System.out.println(ss+ "δ����");
		return ss;
	}
	
	public static void outXls(Workbook readwb,String[][][] table,String[][] writetable, int readSheet) throws RowsExceededException, WriteException {
		// �����Ѿ�������Excel������,�����µĿ�д���Excel������
		jxl.write.WritableWorkbook wwb;
		try {
			wwb = Workbook.createWorkbook(new File(
					NewPath), readwb);
			//��ȡ
			jxl.write.WritableSheet ws = wwb.getSheet(readSheet);
			
			//д��
			for(int i = 2; i < writetable.length ; i++){
				for(int j = 0; j < 30; j++){
					String str = writetable[i][j];
					try {
						double d = Double.parseDouble(str);
						Number number = new Number(j, i, d);
						ws.addCell(number);
					} catch (Exception e) {
						Label label = new Label(j, i, str);
						ws.addCell(label);
					}
				}
				
			}
			
			System.out.println("##�������");
			// д��Excel����
			wwb.write();
			wwb.close();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		
	}

}
