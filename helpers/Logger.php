<?php

namespace lcssoft\report\helpers;;


use yii\console\Application;
use yii\helpers\BaseConsole;

class Logger
{

    /**
     * @param $exception \Exception
     * @param $traceKey
     * @return void
     */
    public static function error($exception, $traceKey = '')
    {
        self::log("message: " . $exception->getMessage(), $traceKey);
        self::log("file: " . $exception->getFile(), $traceKey);
        self::log("line: " . $exception->getLine(), $traceKey);
        self::log("trace: " . $exception->getTraceAsString(), $traceKey);
    }

    public static function log($message, $traceKey = '')
    {
        if (\Yii::$app instanceof Application) {
            BaseConsole::output("[$traceKey] $message");
        } else {
            \Yii::error("[$traceKey] $message");
        }
    }
}