<?php

namespace customit\excelreport;

use Yii;
use yii\queue\closure\Behavior;
use yii\queue\Queue;
use yii\queue\ExecEvent;

class ProgressBehavior extends Behavior
{
    private $jobId;

    public function events()
    {
        return [
            Queue::EVENT_BEFORE_EXEC => function (ExecEvent $event) {
                $this->jobId = $event->id;                
            }
        ];
    }

    public function setProgress($pos, $len)
    {
        $key = __CLASS__ . $this->jobId;
        Yii::$app->cache->set($key, [$pos, $len]);
    }

    public function getProgress($jobId)
    {
        $key = __CLASS__ . $jobId;
        return Yii::$app->cache->get($key) ?: [0, 1];
    }

    public function setManuallyProgress($jobId, $pos, $len) {
        $key = __CLASS__ . $jobId;
        Yii::$app->cache->set($key, [$pos, $len]);
    }
}
