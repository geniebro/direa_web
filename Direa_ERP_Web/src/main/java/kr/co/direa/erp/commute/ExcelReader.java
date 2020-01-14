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
	 * 출퇴근 기록지 엑셀파일을 리스트로 읽어오는 메소드
	 * 
	 * @param filePath
	 * @return
	 * @throws IOException
	 * @throws CustomException 
	 */
	public ArrayList<CommuteVO> read(String filePath) throws IOException, CustomException {
		ArrayList<CommuteVO> voList = new ArrayList<CommuteVO>();
		
		//파일을 읽기위해 엑셀파일을 가져온다
		FileInputStream fis=new FileInputStream(filePath);
		
		HSSFWorkbook workbook=new HSSFWorkbook(fis);
		int rowindex=0;
		int columnindex=0;
		//시트 수 (첫번째에만 존재하므로 0을 준다)
		//만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
		HSSFSheet sheet=workbook.getSheetAt(0);
		
		//행의 수
		int rows=sheet.getPhysicalNumberOfRows();
		for(rowindex=1;rowindex<rows;rowindex++){
			CommuteVO vo = new CommuteVO();
		    //행을 읽는다
		    HSSFRow row=sheet.getRow(rowindex);
		    if(row !=null){
		        //셀의 수
		        int cells=row.getPhysicalNumberOfCells();
		        for(columnindex=0;columnindex<=cells;columnindex++){
		            //셀값을 읽는다
		            HSSFCell cell=row.getCell(columnindex);
		            String value="";
		            //셀이 빈값일경우를 위한 널체크
		            if(cell==null){
		                continue;
		            }else{
		                //타입별로 내용 읽기
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
							throw new CustomException("엑셀 컬럼값이 지원하지 않는 타입입니다. ("+cell.getCellType()+")");
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
