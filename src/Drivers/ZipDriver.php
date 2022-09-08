<?php

namespace AnourValar\Office\Drivers;

class ZipDriver implements DocumentInterface
{
    /**
     * @var array
     */
    protected readonly array $fileSystem;

    /**
     * @var array
     */
    protected array $data = [];

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\LoadInterface::load()
     */
    public function load(string $file, \AnourValar\Office\Format $format): self
    {
        $instance = new static;
        $instance->getFormat($format); // check
        $fileSystem = [];

        $zipArchive = new \ZipArchive;
        $zipArchive->open($file);
        try {
            $count = $zipArchive->numFiles;

            for ($i = 0; $i < $count; $i++) {
                $filename = $zipArchive->getNameIndex($i);
                $content = $zipArchive->getFromName($filename);

                $fileSystem[$filename] = $content;
            }
        } catch (\Throwable $e) {
            $zipArchive->close();
            throw $e;
        }
        $zipArchive->close();

        $instance->fileSystem = $fileSystem;
        return $instance;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SaveInterface::save()
     */
    public function save(string $file, \AnourValar\Office\Format $format): void
    {
        $this->getFormat($format); // check

        $options = \App::make(\ZipStream\Option\Archive::class);
        $zipStream = new \ZipStream\ZipStream(null, $options);
        ob_start();

        try {
            foreach ($this->fileSystem as $filename => $content) {
                if ($content && mb_strtolower(pathinfo($filename, PATHINFO_EXTENSION)) == 'xml') {
                    $content = $this->handleReplace($content);
                }

                $zipStream->addFile($filename, $content);
             }
        } catch (\Throwable $e) {
            $zipStream->finish();
            ob_get_clean();
            throw $e;
        }

        $zipStream->finish();
        file_put_contents($file, ob_get_clean());
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\DocumentInterface::replace()
     */
    public function replace(array $data): self
    {
        foreach ($data as $key => $value) {
            $this->data[$key] = e($value);
        }

        return $this;
    }

    /**
     * @param \AnourValar\Office\Format $format
     * @return string
     */
    protected function getFormat(\AnourValar\Office\Format $format): string
    {
        return match($format) {
            \AnourValar\Office\Format::Docx => 'Docx',
        };
    }

    /**
     * @param string $content
     * @return string
     */
    protected function handleReplace(string $content): string
    {
        foreach ($this->data as $from => $to) {
            $pattern = mb_str_split($from);
            foreach ($pattern as &$patternItem) {
                $patternItem = preg_quote($patternItem);
                $patternItem .= '(\<[^\[]*)?';
            }
            unset($patternItem);
            $pattern = implode('', $pattern);

            $content = preg_replace_callback("#$pattern#Uu", function ($patterns) use ($from, $to) {
                if (strip_tags($patterns[0]) == $from) {
                    return $to;
                }

                return $patterns[0];
            }, $content);
        }

        return $content;
    }
}
