# frozen_string_literal: true

require 'ole/storage'
require 'openssl'
require 'set'
require 'zip'
require 'awesome_print' if $DEBUG
require 'pry' if $DEBUG

module Unxls
  module Log; end # For easy output of values in different formats
  class BitOps; end # For easy binary operations

  module Helpers # Too small yet to move to a separate file
    refine Object do
      # Convert String to StringIO unless it's already StringIO
      # @return [StringIO]
      # @raises [RuntimeError] Unless receiver is a String or a StringIO
      def to_sio
        raise "Cannot convert #{self.class} to StringIO" unless self.kind_of?(StringIO) || self.instance_of?(String)
        self.is_a?(StringIO) ? self : StringIO.new(self)
      end
    end
  end

  module Oshared; end # [MS-OSHARED]: Office Common Data Types and Objects Structures
  module Dtyp; end # [MS-DTYP]: Windows Data Types
  module Offcrypto; end # [MS-OFFCRYPTO]: Office Document Cryptography Structure

  module Biff8 # Group code by BIFF version
    module Constants; end # Large constants

    class WorkbookStream; end # Deals with Workbook Stream (2.1.7.20)

    class Record; end # Recipes for parsing of particular records (2.4)

    module Structure; end # Deals with structures used in records (2.5)
    module ParsedExpressions; end # Extends Structure. Deals with structures used in formulas (2.5.198). Defined in Structure's file for now.

    class Browser; end
  end

  class Parser; end # Opens the compound file and parses the needed storages and streams (2.1.7)

end

require_relative 'version' unless defined?(Unxls::VERSION)
require_relative 'log'
require_relative 'bit_ops'
require_relative '../../ext/bit_ops.so'
require_relative 'biff8/constants'
require_relative 'biff8/workbook_stream'
require_relative 'biff8/record'
require_relative 'biff8/structure'
require_relative 'biff8/browser'
require_relative 'oshared'
require_relative 'dtyp'
require_relative 'offcrypto'
require_relative 'parser'
