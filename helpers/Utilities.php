<?php


namespace lcssoft\report\helpers;;



class Utilities
{
    /**
     * @param $string string
     * @param $index int | array
     * @param $needle string
     * @return string
     * @example
     *  - insertStringIntoStringIndex("ABCDE",2,"+")
     *  => result: "A+BCDE"
     *  - insertStringIntoStringIndex("ABCDE",[2,4],"+")
     *  => result: "A+BC+DE"
     */
    public static function insertStringIntoStringIndex($string,$index,$needle){

        $result = $string;
        $diff = 0;
        if(!is_array($index)){
            $index = [$index];
        }
        foreach ($index as $arrIndex){
            $point = ($arrIndex-1+$diff);
            if($point>strlen($result)){
                continue;
            }
            $tempString= substr($result,0,$point)."$needle";
            $tempString.= substr($result,$point,strlen($string));
            $result = $tempString;
            $diff++;
        }
        return $result;
    }

    public static function getProviderTotal($provider, $fieldName,$condition = ['k' => '','v'=>''])
    {
        $total = 0;

        foreach ($provider as $item) {
            $total += $item[$fieldName];
        }

        return $total;
    }



    public static function createDirectory($directories = [])
    {
        try {
            if (empty($directories)) {
                return false;
            }
            $dirPath = "";
            foreach ($directories as $dir) {
                $dirPath .= DIRECTORY_SEPARATOR . $dir;
                if (!is_dir($dirPath)) {
                    mkdir($dirPath, 0777, true);
                }
            }
            return $dirPath;
        } catch (\Exception $e) {
           throw $e;
        }

        return false;
    }




}