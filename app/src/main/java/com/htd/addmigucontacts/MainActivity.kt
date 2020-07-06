package com.htd.addmigucontacts

import android.Manifest
import android.app.Activity
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
import android.provider.ContactsContract
import android.provider.ContactsContract.RawContacts
import android.provider.ContactsContract.CommonDataKinds.Phone;
import android.provider.ContactsContract.Data;
import android.util.Log
import android.view.View
import androidx.appcompat.app.AppCompatActivity
import io.reactivex.ObservableEmitter
import io.reactivex.Observer
import io.reactivex.disposables.Disposable
import jxl.Workbook

class MainActivity : AppCompatActivity() {

    val data = ArrayList<File>()

    private var dialog: ProgressDialog? = null

    private var imetter:ObservableEmitter<File>? = null ;

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
                        Toast.makeText(this,"请打开相关权限",Toast.LENGTH_SHORT);
                        // 如果权限申请失败，则退出
                        //android.os.Process.killProcess(android.os.Process.myPid())
                    }
                }

//        if (dialog == null) {
//            dialog = ProgressDialog(this)
//            dialog!!.setMessage("添加中...")
//            dialog!!.setCancelable(false)
//        }
        add_btn.setOnClickListener(object :View.OnClickListener{
            override fun onClick(v: View?) {
                pipeWork()
            }
        });
    }

    fun pipeWork(){
        Observable
            .create<File> {
                val intent = Intent()
                    .setType("*/*")
                    .setAction(Intent.ACTION_GET_CONTENT)

                startActivityForResult(Intent.createChooser(intent, "Select a file"), 111)

//                        it.onNext(adapter.data[position] as File)
//                        it.onComplete()
                imetter = it;
            }
            .flatMap {
                Log.i(TAG,""+it.absolutePath)
                val workBook = Workbook.getWorkbook(it)
                val sheet = workBook.getSheet(0) // 获取第一张表格中的数据
                val list = ArrayList<Contacts>()
                // 行数
                for (row in 0 until sheet.rows) {
                    list.add(Contacts(
                        sheet.getCell(0, row).contents, // 第一列是姓名
                        sheet.getCell(1, row).contents  // 第二列是号码
                    ))
                }
                workBook.close()
                return@flatMap Observable.fromIterable(list)
            }
            .subscribeOn(Schedulers.io())
            .doOnSubscribe {
                if (!dialog!!.isShowing && dialog != null) {
                    dialog!!.show()
                }
            }
            .observeOn(AndroidSchedulers.mainThread())
            .subscribe(object : Observer<Contacts> {
                override fun onComplete() {
                    if (dialog!!.isShowing && dialog != null) {
                        dialog!!.dismiss()
                    }
                    toast("添加完毕")
                }

                override fun onSubscribe(d: Disposable) {
                }

                override fun onNext(t: Contacts) {
                    addContact(t)
                }

                override fun onError(e: Throwable) {
                    toast("错误:${e.message}")
                }

            })
    }
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)

        // Selected a file to load
        if ((requestCode == 111) && (resultCode == RESULT_OK)) {
            val selectedFilename = data?.data //The uri with the location of the file
            if (selectedFilename != null) {
                val filenameURIStr = selectedFilename.toString()
                if (filenameURIStr.endsWith(".txt", true)) {
                    val msg = "Chosen file: " + filenameURIStr
                    val toast = Toast.makeText(applicationContext, msg, Toast.LENGTH_SHORT)
                    toast.show()
                    var fil = File(selectedFilename.getPath())
                    if(imetter != null){
                        imetter!!.onNext(fil)
                        imetter!!.onComplete()
                    }
                }
                else {
                    val msg = "The chosen file is not a .txt file!"
                    val toast = Toast.makeText(applicationContext, msg, Toast.LENGTH_LONG)
                    toast.show()
                }
            }
            else {
                val msg = "Null filename data received!"
                val toast = Toast.makeText(applicationContext, msg, Toast.LENGTH_LONG)
                toast.show()
            }
        }
    }
    fun addContact(contacts: Contacts) {
        // 联系人号码可能不止一个，例如 12345678901;12345678901
        val phone = contacts.phone.split(";".toRegex()).dropLastWhile { it.isEmpty() }.toTypedArray()

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
        values.put(StructuredName.GIVEN_NAME, contacts.name)
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


    fun toast(word:String){
        Toast.makeText(this,word,Toast.LENGTH_SHORT);

    }

}
