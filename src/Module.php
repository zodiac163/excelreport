<?php

/**
 * @package   yii2-excel-report
 * @author    Zodiac163 <vm@custom-it.ru>
 * @copyright Copyright &copy; Zodiac163, custom-it.ru
 * @version   0.0.1
 */

namespace customit\excelreport;

class Module extends \yii\base\Module
{
    public function init()
    {
        parent::init();
        $this->registerTranslations();
    }

    public function registerTranslations()
    {
        \Yii::$app->i18n->translations['customit'] = [
            'class' => 'yii\i18n\PhpMessageSource',
            'sourceLanguage' => 'en-US',
            'basePath' => __DIR__ . '/messages',
        ];
    }
}
