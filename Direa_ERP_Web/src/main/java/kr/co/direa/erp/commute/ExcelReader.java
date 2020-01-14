package kr.co.direa.erp.commute;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import kr.co.direa.erp.CustomException;

public class ExcelReader {
	public static void main(String[] args) {
		ExcelReader test = new ExcelReader();
		
		try {
			test.read("D:\\test.xls");
		} catch (Exception e) {
			e.printStackTrace();
		} catch (CustomException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	/**
	 * ����� ����� ���������� ����Ʈ�� �о���� �޼ҵ�
	 * 
	 * @param filePath
	 * @return
	 * @throws IOException
	 * @throws CustomException 
	 */
	public ArrayList<CommuteVO> read(String filePath) throws IOException, CustomException {
		ArrayList<CommuteVO> voList = new ArrayList<CommuteVO>();
		
		//������ �б����� ���������� �����´�
		FileInputStream fis=new FileInputStream(filePath);
		
		HSSFWorkbook workbook=new HSSFWorkbook(fis);
		int rowindex=0;
		int columnindex=0;
		//��Ʈ �� (ù��°���� �����ϹǷ� 0�� �ش�)
		//���� �� ��Ʈ�� �б����ؼ��� FOR���� �ѹ��� �����ش�
		HSSFSheet sheet=workbook.getSheetAt(0);
		
		//���� ��
		int rows=sheet.getPhysicalNumberOfRows();
		for(rowindex=1;rowindex<rows;rowindex++){
			CommuteVO vo = new CommuteVO();
		    //���� �д´�
		    HSSFRow row=sheet.getRow(rowindex);
		    if(row !=null){
		        //���� ��
		        int cells=row.getPhysicalNumberOfCells();
		        for(columnindex=0;columnindex<=cells;columnindex++){
		            //������ �д´�
		            HSSFCell cell=row.getCell(columnindex);
		            String value="";
		            //���� ���ϰ�츦 ���� ��üũ
		            if(cell==null){
		                continue;
		            }else{
		                //Ÿ�Ժ��� ���� �б�
		                switch (cell.getCellType()){
		                case FORMULA:
		                    value=cell.getCellFormula();
		                    break;
		                case NUMERIC:
		                    value=cell.getNumericCellValue()+"";
		                    break;
		                case STRING:
		                    value=cell.getStringCellValue()+"";
		                    break;
		                case BLANK:
		                    value=cell.getBooleanCellValue()+"";
		                    break;
		                case ERROR:
		                    value=cell.getErrorCellValue()+"";
		                    break;
						default:
							throw new CustomException("���� �÷����� �������� �ʴ� Ÿ���Դϴ�. ("+cell.getCellType()+")");
		                }
		            }
		            vo.setField(columnindex, value);
//		            int valueSize = value.length();
//		            String size = "";
//		            if (valueSize < 10) {
//						size = "10";
//					}else {
//						size = "20";
//					}
//		            String format = "%" + size +  "s";
//	            	System.out.printf(format, value);
	            }
		        voList.add(vo);
//	        	System.out.println();
		    }
		}
		return voList;
	}
}
