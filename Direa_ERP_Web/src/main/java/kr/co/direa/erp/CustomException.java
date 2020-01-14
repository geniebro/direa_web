package kr.co.direa.erp;

public class CustomException extends Throwable {
	private String errMsg;
	
	public CustomException(String msg) {
		this.errMsg = msg;
	}
}
