package kr.co.direa.erp.commute;

import kr.co.direa.erp.CustomException;
import lombok.Data;

@Data
public class CommuteVO {
	private String date;
	private String group;
	private String name;
	private String holyday;
	private String startTime;
	private String endTime;
	private String startType;
	private String endType;
	private String lateTime;
	private String earlyWorkTime;
	private String lateWorkTime;
	private String nightWorkTime;
	private String holydayWorkTime;
	private String totalWorkTime;
	private String nomalWorkTime;
	private String goingOut;
	
	public void setField(int num, String value) throws CustomException {
		switch (num) {
		case 0: setDate(value);     		break;
		case 1: setGroup(value);       		break;
		case 2: setName(value);       		break;
		case 3: setHolyday(value);    		break;
		case 4: setStartTime(value);        break;
		case 5: setEndTime(value);       	break;
		case 6: setStartType(value);       	break;
		case 7: setEndType(value);       	break;
		case 8: setLateTime(value);         break;
		case 9: setEarlyWorkTime(value);    break;
		case 10: setLateWorkTime(value);     break;
		case 11: setNightWorkTime(value);   break;
		case 12: setHolydayWorkTime(value); break;
		case 13: setTotalWorkTime(value);   break;
		case 14: setNomalWorkTime(value);   break;
		case 15: setGoingOut(value);        break;

		default:
			throw new CustomException("엑셀파일의 컬럼 수가 정의보다 많습니다. ("+num+")("+value+")");
		}
	}
}
