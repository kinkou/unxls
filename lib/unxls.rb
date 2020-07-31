# frozen_string_literal: true

require_relative 'unxls/map'

module Unxls
  # @param path [String, Pathname]
  # @param settings [Hash]
  # @return [Hash]
  def self.parse(path, settings = {})
    file = File.open(path, 'rb')
    Unxls::Parser.new(file, settings).parse
  ensure
    file.close if file
  end

end
