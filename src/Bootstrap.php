<?php

namespace customit\excelreport;

use Yii;
use yii\base\BootstrapInterface;
class Bootstrap implements BootstrapInterface{
    public function bootstrap($app)
    {
        $app->setModule('excelreport', 'customit\excelreport\Module');
        Yii::$app->getModule('excelreport')->registerTranslations();
    }
}