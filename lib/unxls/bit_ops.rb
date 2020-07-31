# frozen_string_literal: true

module Unxls
  class BitOps

    # @param bits [Integer]
    def initialize(bits)
      @bits = bits
    end

    # @example
    #   BitOps.new(0b010).set_at?(1) # -> true
    # @param index [Integer] 0-based
    # @return [true, false, nil]
    def set_at?(index)
      return nil if index < 0
      @bits[index] == 1
    end

    # @example
    #   BitOps.new(0b1110111).value_at(2..4) # -> 0b101
    #   BitOps.new(0b1110111).value_at(2) # -> 0b1
    # @param range [Range, Integer] 0-based
    # @return [Integer, nil]
    def value_at(range)
      range = (range..range) if range.is_a?(Integer)
      bits_in_result = range.size
      offset = range.min
      return nil if !offset || offset < 0
      mask = make_mask(bits_in_result, offset)
      (@bits & mask) >> offset
    end

    # @example
    #   BitOps.new(0b1100101).reverse # -> 0b1010011
    # @return [Integer]
    def reverse
      number = @bits
      result = 0
      while number > 0 do
        result = result << 1
        result = result | (number & 1)
        number = number >> 1
      end
      result
    end

    # @example
    #   .ror(0b11110000, 2, 8) # -> 0b11000011
    # @param bitsize [Integer]
    # @param steps [Integer]
    # @return [Integer]
    def rol(bitsize, steps)
      rotate(bitsize, steps, :left)
    end

    # @example
    #   .ror(0b11110000, 2, 8) # -> 0b00111100
    # @param bitsize [Integer]
    # @param steps [Integer]
    # @return [Integer]
    def ror(bitsize, steps)
      rotate(bitsize, steps, :right)
    end

    private

    # @param bitsize [Integer]
    # @param steps [Integer]
    # @param direction [Symbol] :left, :right
    # @return [Integer]
    def rotate(bitsize, steps, direction)
      case direction
      when :left then (@bits << steps) | (@bits >> (bitsize - steps))
      when :right then (@bits >> steps) | (@bits << (bitsize - steps))
      else raise "Unexpected rotate direction #{direction}"
      end & make_mask(bitsize, 0)
    end

    # @example
    #   make_mask(3) # -> 0b111
    #   make_mask(3, 2) # -> 0b11100
    # @param length [Integer] 0-based
    # @param offset [Integer] 0-based
    # @return [Integer]
    def make_mask(length, offset)
      true_bits = 2 ** length - 1
      true_bits << offset
    end

  end
end