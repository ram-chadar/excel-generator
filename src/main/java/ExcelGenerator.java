import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelGenerator {

	public static void main(String[] args) {
		System.out.println("Application Started....");
		System.out.println("***********");
		AddName nameObj = new AddName();
		String name=nameObj.addName();
		System.out.println(name);
		String[] headers = { "Question", "Option A", "Option B", "Option C", "Option D", "Answer" };

		// Full Data for 40 Questions
		String[][] data = {
				{ "What is the purpose of Angular directives?", "To define components",
						"To add behavior to an existing DOM element", "To handle services", "To manage routing", "B" },
				{ "Which of the following is a structural directive in Angular?", "ngIf", "ngClass", "ngStyle",
						"ngBind", "A" },
				{ "What symbol is used to apply a structural directive in Angular?", "#", "*", "@", "&", "B" },
				{ "Which directive is used to repeat a part of the DOM tree for each item in a list?", "*ngIf",
						"*ngFor", "ngBind", "*ngSwitch", "B" },
				{ "How can you apply conditional styles using Angular directives?", "[ngIf]", "[ngClass]", "[ngFor]",
						"[ngSwitch]", "B" },
				{ "Which of the following Angular directives is used to toggle DOM elements?", "ngFor", "ngIf",
						"ngClass", "ngSwitch", "B" },
				{ "What is the purpose of the ng-template directive?", "To create a reusable component",
						"To define a template that can be used conditionally", "To apply styles", "To create a form",
						"B" },
				{ "Which of the following is NOT a built-in directive in Angular?", "ngModel", "ngSwitch", "ngClass",
						"ngFilter", "D" },
				{ "Which directive can be used to conditionally switch between different templates?", "*ngSwitch",
						"*ngIf", "*ngClass", "*ngFor", "A" },
				{ "How do you bind a custom directive to an attribute in Angular?", "@Directive",
						"[ ] (square brackets)", "{{ }} (double curly braces)", "< > (angle brackets)", "B" },
				{ "What is the primary purpose of Angular pipes?", "To control routing",
						"To format data in the template", "To manage forms", "To create components", "B" },
				{ "Which of the following is a built-in pipe in Angular?", "CurrencyPipe", "ngFilterPipe",
						"ngStylePipe", "FilterPipe", "A" },
				{ "What pipe can be used to transform an array into a sorted array?", "json", "lowercase", "orderBy",
						"There is no such pipe in Angular", "D" },
				{ "Which pipe is used to convert a string to uppercase in Angular?", "lowercase", "uppercase",
						"titlecase", "capitalize", "B" },
				{ "How do you use a custom pipe in Angular?", "Add it to NgModule declarations",
						"Use @Pipe decorator and add it to declarations", "Register in the component file only",
						"No need to register", "B" },
				{ "Which pipe is used to format numbers in Angular?", "DatePipe", "NumberPipe", "CurrencyPipe",
						"DecimalPipe", "D" },
				{ "What is the output of {{ 'hello' | uppercase }}?", "Hello", "HELLO", "hello", "HeLLo", "B" },
				{ "Which pipe would you use to display a date in Angular?", "DatePipe", "CurrencyPipe", "TimePipe",
						"StringPipe", "A" },
				{ "Which pipe converts an object into a JSON string?", "CurrencyPipe", "JsonPipe", "StringPipe",
						"ObjectPipe", "B" },
				{ "How would you use the pipe decimal with 2 decimal points?", "{{ value | decimal:'2.2' }}",
						"{{ value | number:'1.2-2' }}", "{{ value | decimal:'1.2-2' }}", "{{ value | decimal }}", "B" },
				{ "How do you create a component in Angular using the CLI?", "ng new component-name",
						"ng create component-name", "ng generate component component-name",
						"ng add component component-name", "C" },
				{ "What decorator is used to define an Angular component?", "@Injectable", "@Pipe", "@Component",
						"@Directive", "C" },
				{ "Which of the following is mandatory in a component decorator?", "template", "selector", "styles",
						"providers", "B" },
				{ "What is the purpose of the selector in an Angular component?", "To apply CSS styles",
						"To inject services", "To define how the component will be identified in templates",
						"To manage component states", "C" },
				{ "How can you pass data from a parent to a child component?", "@Inject", "@Input", "@Output",
						"@ViewChild", "B" },
				{ "How do you capture events emitted from a child component?", "@Input", "@Output", "EventEmitter",
						"Both B and C", "D" },
				{ "Which lifecycle hook is called when the component is first initialized?", "ngOnInit",
						"ngAfterViewInit", "ngOnChanges", "ngDoCheck", "A" },
				{ "Which Angular lifecycle hook is called after the component's view has been fully initialized?",
						"ngOnChanges", "ngAfterViewInit", "ngOnDestroy", "ngOnInit", "B" },
				{ "In which method should you perform component cleanup to prevent memory leaks?", "ngAfterViewInit",
						"ngOnChanges", "ngOnDestroy", "ngDoCheck", "C" },
				{ "What is the correct way to add CSS styles specifically to a component?",
						"In the component’s styles property", "In the global styles.css",
						"In the app.component.css file", "Using inline styles in the HTML template", "A" } };

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Angular Questions Sheet");

		// Create the header row
		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
		}

		// Create data rows
		int rowNum = 1;
		for (Object[] rowData : data) {
			Row row = sheet.createRow(rowNum++);
			for (int colNum = 0; colNum < rowData.length; colNum++) {
				Cell cell = row.createCell(colNum);
				if (rowData[colNum] instanceof String) {
					cell.setCellValue((String) rowData[colNum]);
				} else if (rowData[colNum] instanceof Integer) {
					cell.setCellValue((Integer) rowData[colNum]);
				}
			}
		}

		// Auto-size columns for readability
		for (int i = 0; i < headers.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		try (FileOutputStream fileOut = new FileOutputStream("Angular - Level 2.xlsx")) {
			workbook.write(fileOut);

		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				workbook.close();
			} catch (IOException e) {

				e.printStackTrace();
			}
		}

		// Closing the workbook
		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Excel file created successfully  Done!");
		System.out.println("**********");
		
		
	}
}
