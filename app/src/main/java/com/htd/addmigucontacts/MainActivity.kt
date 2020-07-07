package com.htd.addmigucontacts

import android.Manifest
import android.app.ProgressDialog
import android.content.ContentValues
import android.os.Bundle
import android.widget.Toast
import com.tbruyelle.rxpermissions2.RxPermissions
import io.reactivex.Observable
import io.reactivex.android.schedulers.AndroidSchedulers
import io.reactivex.schedulers.Schedulers
import kotlinx.android.synthetic.main.activity_main.*
import java.io.File
import android.provider.ContactsContract.CommonDataKinds.StructuredName
import android.content.ContentUris
import android.content.Intent
import android.os.Environment
import android.provider.ContactsContract
import android.provider.ContactsContract.RawContacts
import android.provider.ContactsContract.CommonDataKinds.Phone;
import android.provider.ContactsContract.Data;
import android.util.Log
import android.view.View
import androidx.appcompat.app.AppCompatActivity
import io.reactivex.Observer
import io.reactivex.disposables.Disposable
import org.apache.poi.hssf.usermodel.HSSFDateUtil
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.text.SimpleDateFormat
import kotlin.reflect.typeOf
import kotlin.Number as Number1

class MainActivity : AppCompatActivity() {

    private var dialog: ProgressDialog? = null

    private var TAG = "zhu";

    init {
        //adapter = GeneralAdapter(R.layout.item_excel, data, BR.bean)
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        // 检查权限
        val rxPermissions = RxPermissions(this)
        rxPermissions
            .request(
                Manifest.permission.READ_EXTERNAL_STORAGE,
                Manifest.permission.WRITE_EXTERNAL_STORAGE,
                Manifest.permission.READ_CONTACTS,
                Manifest.permission.WRITE_CONTACTS
            )
            .subscribe {
                if (!it) {
                    Toast.makeText(this, "请打开相关权限", Toast.LENGTH_SHORT);
                    // 如果权限申请失败，则退出
                    //android.os.Process.killProcess(android.os.Process.myPid())
                }
            }

//        if (dialog == null) {
//            dialog = ProgressDialog(this)
//            dialog!!.setMessage("添加中...")
//            dialog!!.setCancelable(false)
//        }
        //打开系统文件管理器
        //https://stackoom.com/question/404dO/Android-Kotlin-%E8%8E%B7%E5%8F%96FileNotFoundException%E5%B9%B6%E4%BB%8E%E6%96%87%E4%BB%B6%E9%80%89%E6%8B%A9%E5%99%A8%E4%B8%AD%E9%80%89%E6%8B%A9%E6%96%87%E4%BB%B6%E5%90%8D
        add_btn.setOnClickListener(object : View.OnClickListener {
            override fun onClick(v: View?) {
//                rxTest();
                Log.i(TAG,"按钮点击了")
                val intent = Intent()
                    .setType("*/*")
                    // intent.setDataAndType(path, "application/excel");
                    .setAction(Intent.ACTION_GET_CONTENT)

                startActivityForResult(Intent.createChooser(intent, "Select a file"), 111)
            }
        });
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        Log.i(TAG, "onActivityResult()");
        // Selected a file to load
        if ((requestCode == 111) && (resultCode == RESULT_OK)) {
            val selectedFilename = data?.data //The uri with the location of the file
            if (selectedFilename != null) {
                val filenameURIStr = selectedFilename.toString()
                if (filenameURIStr.endsWith(".xlsx", true)) {
                    val msg = "Chosen file: " + filenameURIStr
                    val toast = Toast.makeText(applicationContext, msg, Toast.LENGTH_SHORT)
                    toast.show()
                    var path =
                        Environment.getExternalStorageDirectory().toString() + "/360/migu2.xlsx"
                    //var path2 = selectedFilename.path
                    //   /external_files/360/migu.xlsx
                    //var lastPath = path+path2!!.substring(12)
                    //Log.i(TAG,"实际路径"+path)
                    //Log.i(TAG,"应该路径"+path2)
                    //Log.i(TAG,"最后路径"+lastPath)

                    var fil = File(path)
//                        imetter!!.onNext(fil)
//                        imetter!!.onComplete()
                    //pipeWork(fil)
                    readXslx(fil)
                } else {
                    val msg = "The chosen file is not a .txt file!"
                    val toast = Toast.makeText(applicationContext, msg, Toast.LENGTH_LONG)
                    toast.show()
                }
            } else {
                val msg = "Null filename data received!"
                val toast = Toast.makeText(applicationContext, msg, Toast.LENGTH_LONG)
                toast.show()
            }
        }
    }

    fun rxTest() {
        Observable.create<String> {
            it.onNext("Hello Rx!")
            it.onComplete()
        }
            .subscribe(object : Observer<String> {
                override fun onComplete() {

                    Log.i(TAG, "添加完毕")
                }

                override fun onSubscribe(d: Disposable) {

                }

                override fun onNext(t: String) {
                    Log.i(TAG, "onNext")
                }

                override fun onError(e: Throwable) {
                    Log.i(TAG, "错误 = " + e.message)
                }

            })
    }

    //https://www.jianshu.com/p/9bf1f7f7b642
    fun readXslx(file: File) {
        Log.i(TAG, "readXslx()")
        Observable
            .create<String> {
                val list = ArrayList<Contacts>()
                val workbook = XSSFWorkbook(file)
                val sheet = workbook.getSheetAt(0)
                val rowsCount = sheet.getPhysicalNumberOfRows()
                Log.i(TAG, "总行数： " + rowsCount)
                val formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator()

                for (r in 0 until rowsCount) {
                    Log.i(TAG,"- - - - - - - - - - - - - - - 第${r}行 - - - - - - - - - - - - - - - ")
                    var people = Contacts();
                    val row = sheet.getRow(r)
//                    if (r == 0) {
//                        Log.i(TAG, "row.toString() => "+row.toString());
//                    }
                    val cellsCount = row.getPhysicalNumberOfCells()
                    for (c in 0..cellsCount) {
                        var value = getCellAsString(row, c, formulaEvaluator)

                        val cellInfo = "第几行r:$r;   第几列c:$c;    值v:$value"
                        if(c==14){ //名字
                            people.name = value
                        }
                        if(c==15){ //地址
                            if(value.length>6){
                                people.local = value.substring(0,6)
                            }else{
                                people.local = value
                            }
                        }

                        if(c==16){ //电话
                            //var a = typeof value;
                            //Log.i(TAG, ""+a)
                            people.phone = value
                            Log.i(TAG,""+people.name+"  "+people.local+"  "+people.phone)
                            if(people.phone.contains("1")){
                                addContact(people);
                            }
                        }
                    }
                }
                workbook.close()

                it.onNext("some")
                it.onComplete()
            }
            .subscribeOn(Schedulers.io())
            .doOnSubscribe {
                runOnUiThread(Runnable {
                    if (dialog == null) {
                        dialog = ProgressDialog(this)
                        dialog!!.setMessage(".xlsx 文件读取中...")
                        dialog!!.setCancelable(false)
                    }
                    if (!dialog!!.isShowing && dialog != null) {
                        dialog!!.show()
                    }
                })
            }
            .observeOn(AndroidSchedulers.mainThread())
            .subscribe(object : Observer<String> {
                override fun onComplete() {
                    runOnUiThread(Runnable {
                        if (dialog!!.isShowing && dialog != null) {
                            dialog!!.dismiss()
                        }
                    })
                    Log.i(TAG, "添加完毕")
                }

                override fun onSubscribe(d: Disposable) {

                }

                override fun onNext(t: String) {
                    Log.i(TAG, "onNext")
                }

                override fun onError(e: Throwable) {
                    Log.i(TAG, "错误 = " + e.message)
                }

            })
    }

    fun addContact(contacts: Contacts) {
        Log.i(TAG,"addContact() => "+contacts.toString())

        var title = contacts.name + "-"+contacts.local
        // 联系人号码可能不止一个，例如 12345678901;12345678901
        val phone =
            contacts.phone.split(";".toRegex()).dropLastWhile { it.isEmpty() }.toTypedArray()

        // 创建一个空的ContentValues
        val values = ContentValues()
        // 向RawContacts.CONTENT_URI空值插入，
        // 先获取Android系统返回的rawContactId
        // 后面要基于此id插入值
        val rawContactUri = contentResolver.insert(RawContacts.CONTENT_URI, values)
        val rawContactId = ContentUris.parseId(rawContactUri)
        values.clear()

        values.put(Data.RAW_CONTACT_ID, rawContactId)
        // 内容类型
        values.put(Data.MIMETYPE, StructuredName.CONTENT_ITEM_TYPE)
        // 联系人名字
        values.put(StructuredName.GIVEN_NAME, title)
        // 向联系人URI添加联系人名字
        contentResolver.insert(ContactsContract.Data.CONTENT_URI, values)
        values.clear()

        values.put(Data.RAW_CONTACT_ID, rawContactId)
        values.put(Data.MIMETYPE, Phone.CONTENT_ITEM_TYPE)
        // 联系人的电话号码
        values.put(Phone.NUMBER, phone[0])
        // 电话类型
        values.put(Phone.TYPE, Phone.TYPE_MOBILE)
        // 向联系人电话号码URI添加电话号码
        contentResolver.insert(Data.CONTENT_URI, values)
        values.clear()

        // 当联系人存在多个号码，第二个号码存在工作电话
        if (phone.size > 1) {
            values.put(Data.RAW_CONTACT_ID, rawContactId)
            values.put(Data.MIMETYPE, Phone.CONTENT_ITEM_TYPE)
            // 联系人的工作电话号码
            values.put(Phone.NUMBER, phone[1])
            // 电话类型
            values.put(Phone.TYPE, Phone.TYPE_WORK_MOBILE)
            // 向联系人电话号码URI添加电话号码
            contentResolver.insert(Data.CONTENT_URI, values)
            values.clear()
        }
    }


    fun toast(word: String) {
        Toast.makeText(this, word, Toast.LENGTH_SHORT);
    }

    protected fun getCellAsString(row: Row, c: Int, formulaEvaluator: FormulaEvaluator): String {
        var value = ""
        try {
            val cell = row.getCell(c)
            val cellValue = formulaEvaluator.evaluate(cell)
            if(cellValue == null){
                value =  "null"
            }else{
                when (cellValue.cellType) {
                    Cell.CELL_TYPE_BOOLEAN -> value = "" + cellValue.booleanValue
                    Cell.CELL_TYPE_NUMERIC -> {
                        val numericValue = cellValue.numberValue
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            val date = cellValue.numberValue
                            val formatter = SimpleDateFormat("dd/MM/yy")
                            value = formatter.format(HSSFDateUtil.getJavaDate(date))
                        } else {
                            value = "" + numericValue
                        }
                    }
                    Cell.CELL_TYPE_STRING -> value = "" + cellValue.stringValue
                }
            }
        } catch (e: NullPointerException) {
            /* proper error handling should be here */
            Log.i(TAG,"getCellAsString（）错误 "+e.toString());
            value =  "kong"
        }
        return value
    }



}

