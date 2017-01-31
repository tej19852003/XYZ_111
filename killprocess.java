package Utilities;

import java.io.IOException;

public class killprocess {

	public static void main(String[] args) {

		try {
		    Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
		} catch (IOException e) {
		    e.printStackTrace();
		}
	}
}
