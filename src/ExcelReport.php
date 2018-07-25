<?php

namespace customit\excelreport;

use Yii;
use yii\base\Widget;

class ExcelReport extends Widget
{
    public $columns;
    public $stripHtml;
    public $searchClass;
    public $searchMethod = 'search';
    public $searchParams = [];

    private $formOptions;
    private $id = "excelReportFrom";

    public function init()
    {
        parent::init();
        $this->formOptions = ['id' => $this->id,];
    }

    public function run()
    {
        if (isset($_POST['excelReportAction']) || Yii::$app->session->has('excel-report-progress')) {
            if (Yii::$app->session->has('excel-report-progress')) {
                $data = unserialize(Yii::$app->session->get('excel-report-progress'));
                $fileName = $data['fileName'];
                $id = $data['queueid'];
            } else {
                $fileName = Yii::$app->security->generateRandomString();
                $searchArray = array_merge([Yii::$app->request->queryParams], $this->searchParams);
                $id = Yii::$app->queue->push(new ExcelReportQueue([
                    'columns' => $this->columns,
                    'stripHtml' => $this->stripHtml,
                    'fileName' => $fileName,
                    'searchClass' => $this->searchClass,
                    'searchMethod' => $this->searchMethod,
                    'searchParams' => $searchArray,
                ]));

                Yii::$app->session->set('excel-report-progress', serialize(['fileName' => $fileName, 'queueid' => $id]));
            }

            return $this->render('_progress', [
                'options' => $this->formOptions,
                'queueId' => $id,
                'file' => Yii::$app->basePath . '/runtime/export/' . $fileName . '.xlsx',
            ]);
        } else {
            return $this->render('_form', ['options' => $this->formOptions]);
        }
    }
}