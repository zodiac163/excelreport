<?php

namespace customit\excelreport;

use yii\base\BaseObject;
use yii\queue\RetryableJobInterface;

class ExcelReportQueue extends BaseObject implements RetryableJobInterface {

    public $columns;
    public $stripHtml;
    public $fileName;
    public $searchClass;
    public $searchMethod;
    public $searchParams;

    public function execute($queue) {
        $m = new ExcelReportModel(
            $this->columns,
            $queue,
            $this->fileName,
            $this->searchClass,
            $this->searchMethod,
            $this->searchParams
        );
        $m->start();
    }

    public function getTtr()
    {
        return 5 * 60;
    }

    public function canRetry($attempt, $error)
    {
        return false;
    }
}
