<?php

namespace customit\excelreport;

use Yii;
use SuperClosure\Serializer;

class ExcelReportHelper {
    
    public static function closureDetect($arr) {
        $serializer = new Serializer();
        foreach ($arr as $key=>&$value) {
            if (is_array($value)) {
                $value = self::closureDetect($value);
            } elseif (is_object($value) && self::is_closure($value)) {
                $value = $serializer->serialize($value);
            }
        }
        
        return $arr;
    }
    
    public static function reverseClosureDetect($arr) {
        $serializer = new Serializer();
        foreach ($arr as $key=>&$value) {
            if (is_array($value)) {
                $value = self::reverseClosureDetect($value);
            } elseif (is_string($value) && strpos($value, "SuperClosure")) {
                $value = $serializer->unserialize($value);                
            }
        }
        
        return $arr;
    }

    public static function is_closure($t) {        
        return $t instanceof \Closure;
    }
}

