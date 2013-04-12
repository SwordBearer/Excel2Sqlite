package xmu.swordbearer.excel2sqlite;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import android.app.Activity;
import android.content.res.AssetManager;
import android.database.Cursor;
import android.os.Bundle;
import android.os.Handler;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.ListView;

public class MainActivity extends Activity {
	String TAG = "MainActivity";

	private ListView listView;
	private Button btnImport;

	private Handler handler;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		// 使用HSSF创建excel文档
		Excel2SQLiteHelper.createExcel(this);

		handler = new Handler();

		btnImport = (Button) findViewById(R.id.btn_import_excel);
		listView = (ListView) findViewById(R.id.listview);

		btnImport.setOnClickListener(new View.OnClickListener() {
			public void onClick(View v) {
				new Thread(new Runnable() {
					public void run() {
						importExcel2Sqlite();
					}
				}).start();
			}
		});
	}

	private void importExcel2Sqlite() {
		AssetManager am = this.getAssets();
		InputStream inStream;
		Workbook wb = null;
		try {
			// 读取.xls文档:放在assets文件夹中
			inStream = am.open("students.xls");
			// HSSF
			wb = new HSSFWorkbook(inStream);
			inStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		if (wb == null) {
			Log.e(TAG, "从assets 读取Excel文档出错");
			return;
		}
		// DBAdapter 封装了SQLite的所有操作
		DBAdapter dbAdapter = new DBAdapter(this);
		Sheet sheet1 = wb.getSheetAt(0);// 第一张表单
		Sheet sheet2 = wb.getSheetAt(1);
		if (sheet1 == null) {
			return;
		}
		dbAdapter.open();
		//
		Excel2SQLiteHelper.insertExcelToSqlite(dbAdapter, sheet1);
		Excel2SQLiteHelper.insertExcelToSqlite(dbAdapter, sheet2);
		dbAdapter.close();
		Log.e(TAG, "数据插入成功");
		handler.post(new Runnable() {
			public void run() {
				displayData();
			}
		});
	}

	private void displayData() {
		DBAdapter dbAdapter = new DBAdapter(this);
		dbAdapter.open();
		Cursor allDataCursor = dbAdapter.getAllRow(DBAdapter.STU_TABLE);
		MyCursorAdapter adapter = new MyCursorAdapter(this, allDataCursor);
		listView.setAdapter(adapter);
	}
}
