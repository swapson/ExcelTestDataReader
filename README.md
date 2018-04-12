# Excel Test Case Reader
- Simple excel test case reader
- Worksheet contains 3 sheet: TestCases, Data, Reference
- Sample TestCases sheet

| Test Case Name | Description | Run Mode |
| -------------- | ----------- | -------- |
| TestCase1 | Test case 1 | Y |
| TestCase2 | Test case 2 | N |
| TestCase3 | Test case 3 | Y |

- Sample Data sheet

| TestCase1 START |
| --------------- |  

| Run Mode | Field 1 | Field 2 | Field 3 | Field 4 | Field 5 |
| -------- | ------- | ------- | ------- | ------- | ------- |
| Y | 11 | 12.2 | 12-Apr-18 | TRUE | tc1r1c5 |
| Y | **${userName}** | **${myDate}** | tc1r2c3 | tc1r2c4 | tc1r2c5 |
| N | tc1r3c1 | tc1r3c2 | tc1r3c3 | tc1r3c4 | tc1r3c5 |

| TestCase1 END |
| --------------- |


    - Begining of test case is marked by TEST_CASE_NAME START
    - End of test case is marked by TEST_CASE_NAME END
    - ${userName} & ${myDate} : will be substituded from Reference sheet.
  
- Sample Reference sheet

| Place Holder | Value |
| ------------ | ----- |
| ${userName} | swapnil |
| ${myDate} | 1-Apr-18 |


- Sample Code
```java
public class Main {
	private static final String FILE_PATH = "./src/main/resources/TestCaseData1.xlsx";

	public static void main(String[] args) {
		ExcelReader er;
		try {
			er = new ExcelReader(FILE_PATH);
			Map<String,List<Map<String,String>>> allTestData = er.readAllTestData(false);
			System.out.println("=== Reference Data ===");
			System.out.println(ExcelReader.REFERENCE_DATA);
			
			System.out.println("=== Test Data ===");
			System.out.println(allTestData);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}
}

```

output

=== Reference Data ===

{${myDate}=Sun Apr 01 00:00:00 IST 2018, ${userName}=swapnil}

=== Test Data ===

{TestCase1=[{Run Mode=Y, Field 1=11.0, Field 2=12.2, Field 3=Thu Apr 12 00:00:00 IST 2018, Field 4=true, Field 5=tc1r1c5}, {Run Mode=Y, Field 1=swapnil, Field 2=Sun Apr 01 00:00:00 IST 2018, Field 3=tc1r2c3, Field 4=tc1r2c4, Field 5=tc1r2c5}], TestCase3=[{Run Mode=Y, Col 1=tc3r2c1, Col 2=tc3r2c2, Col 3=tc3r2c3, Col 4=tc3r2c4, Col 5=tc3r2c5, Col 6=tc3r2c6}, {Run Mode=Y, Col 1=tc3r3c1, Col 2=tc3r3c2, Col 3=tc3r3c3, Col 4=tc3r3c4, Col 5=tc3r3c5, Col 6=tc3r3c6}]}
