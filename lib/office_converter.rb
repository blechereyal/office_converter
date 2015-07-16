require "office_converter/version"
require "uri"
require "net/http"
require "tmpdir"
require "spoon"

module OfficeConverter
=begin
  source -> .docx file to be converted
  target_dir -> directory where file will pe persisted
=end
  def self.convert(source, target_dir, soffice_command = nil, convert_to = nil)
    Converter.new(source, target_dir, soffice_command, convert_to).convert
  end

  class Converter
    attr_accessor :soffice_command

    def initialize(source, target_dir, soffice_command = nil, convert_to = nil)
      @source = source
      @target_dir = target_dir
      @soffice_command = soffice_command
      @convert_to = convert_to || "pdf"
      determine_soffice_command
      check_source_type

      unless @soffice_command && File.exists?(@soffice_command)
        raise IOError, "Can't find Libreoffice or Openoffice executable."
      end
    end

    def convert
      # puts ["HOME=/tmp", @soffice_command, "--headless", "--convert-to", @convert_to, @source, "--outdir", @target_dir].join(' ')
      orig_stdout = $stdout.clone
      $stdout.reopen File.new('/dev/null', 'w')
      pid = Spoon.spawnp("/usr/bin/env",@soffice_command, "--headless", "--convert-to", @convert_to, @source, "--outdir", @target_dir)
      Process.waitpid(pid)
      $stdout.reopen orig_stdout
    end

    private

    def determine_soffice_command
      unless @soffice_command
        @soffice_command ||= which("soffice")
        @soffice_command ||= which("soffice.bin")
      end
    end

    def which(cmd)
      exts = ENV['PATHEXT'] ? ENV['PATHEXT'].split(';') : ['']
      ENV['PATH'].split(File::PATH_SEPARATOR).each do |path|
        exts.each do |ext|
          exe = File.join(path, "#{cmd}#{ext}")
          return exe if File.executable? exe
        end
      end

      return nil
    end

    def check_source_type
      return if File.exists?(@source) && !File.directory?(@source) #file
      raise IOError, "Source (#{@source}) is neither a file nor an URL."
    end
  end
end
