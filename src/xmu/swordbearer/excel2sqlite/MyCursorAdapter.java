package xmu.swordbearer.excel2sqlite;

import android.content.Context;
import android.database.Cursor;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.CursorAdapter;
import android.widget.TextView;

public class MyCursorAdapter extends CursorAdapter {
	private LayoutInflater inflater;

	public MyCursorAdapter(Context context, Cursor c) {
		super(context, c);
		inflater = LayoutInflater.from(context);
	}

	@Override
	public View newView(Context context, Cursor cursor, ViewGroup parent) {
		return inflater.inflate(R.layout.layout_listview_item, parent, false);
	}

	@Override
	public void bindView(View view, Context context, Cursor cursor) {
		TextView name = (TextView) view.findViewById(R.id.textview_name);
		TextView score = (TextView) view.findViewById(R.id.textview_score);
		TextView age = (TextView) view.findViewById(R.id.textview_age);
		TextView birth = (TextView) view.findViewById(R.id.textview_birth);
		name.setText(cursor.getString(1));// name
		score.setText(cursor.getString(2));// score
		age.setText(cursor.getString(3));// age
		birth.setText(cursor.getString(4));// birth
	}
}
