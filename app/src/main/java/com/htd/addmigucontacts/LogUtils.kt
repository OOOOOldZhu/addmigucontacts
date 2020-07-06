package com.htd.addmigucontacts

import android.content.Context
import android.util.Log

/**
 * Created by zhu on 2018-01-30.
 */
object LogUtils {

    private val TAG = "zhu"

    fun i(context: Context, msg: String) {
        Log.i(TAG, "${context.javaClass.name} -> $msg")
    }
}