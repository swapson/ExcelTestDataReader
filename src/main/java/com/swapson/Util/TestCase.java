package com.swapson.Util;

/**
 * 
 * @author Swapnil Sonar
 *
 */
public class TestCase {
	private String testCaseName;
	private String testCaseDesc;
	private String runMode;

	public TestCase(String testCaseName, String testCaseDesc, String runMode) {
		this.testCaseName = testCaseName;
		this.testCaseDesc = testCaseDesc;
		this.runMode = runMode;
	}

	public String getTestCaseName() {
		return testCaseName;
	}

	public void setTestCaseName(String testCaseName) {
		this.testCaseName = testCaseName;
	}

	public String getTestCaseDesc() {
		return testCaseDesc;
	}

	public void setTestCaseDesc(String testCaseDesc) {
		this.testCaseDesc = testCaseDesc;
	}

	public String getRunMode() {
		return runMode;
	}

	public void setRunMode(String runMode) {
		this.runMode = runMode;
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("TestCase [testCaseName=");
		builder.append(testCaseName);
		builder.append(", testCaseDesc=");
		builder.append(testCaseDesc);
		builder.append(", runMode=");
		builder.append(runMode);
		builder.append("]\n");
		return builder.toString();
	}

}
