# frozen_string_literal: true

require_relative 'lib/unxls/version'

Gem::Specification.new do |gem|
  gem.authors       = ['Sergey Konotopov']
  gem.email         = 'werk@mail.ru'
  gem.summary       = 'Parser for Microsoft Excel .xls files'
  gem.homepage      = 'https://github.com/kinkou/unxls'
  gem.license       = 'MIT'

  gem.name          = 'unxls'
  gem.files         = Dir['lib/**/*.rb'] + Dir['ext/**/*.rb']
  gem.extensions    = ['ext/extconf.rb']
  gem.test_files    = Dir['spec/**/*.rb']
  gem.require_paths = ['lib']
  gem.version       = Unxls::VERSION

  gem.required_ruby_version = '~> 2.1'

  gem.add_runtime_dependency 'ruby-ole'
  gem.add_runtime_dependency 'rubyzip'

  gem.add_development_dependency 'bundler'
  gem.add_development_dependency 'rake'
  gem.add_development_dependency 'rspec'
  gem.add_development_dependency 'pry'
  gem.add_development_dependency 'awesome_print'
end