<?php

namespace AnourValar\Office\Traits;

trait Parser
{
    /**
     * @param array $data
     * @param string $prefix
     * @return array
     */
    protected function dot(array $data, string $prefix = ''): array
    {
        $result = [];

        foreach ($data as $key => $value) {
            if (is_array($value)) {
                $result = array_replace($result, $this->dot($value, $prefix.$key.'.'));
            } else {
                $result[$prefix.$key] = $value;
            }
        }

        return $result;
    }
}
