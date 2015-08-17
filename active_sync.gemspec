# -*- encoding: utf-8 -*-

$:.push File.expand_path('../lib', __FILE__)
require 'active_sync/version'

Gem::Specification.new do |s|
  s.name = "active-sync"
  s.summary = "A Ruby interface to Active synce."
  s.description = ""
  s.version = Gmail::VERSION
  s.platform = Gem::Platform::RUBY
  s.authors = ["Nicolas Marlier"]
  s.homepage = "http://github.com/nmarlier/active-sync"
  s.licenses = ['MIT']

  # runtime dependencies
  s.add_dependency "yaml"
  s.add_dependency "json"
  s.add_dependency "httpclient"
  
  # development dependencies
  s.add_development_dependency "rake"
  s.add_development_dependency "test-unit"
  s.add_development_dependency('mocha', '~> 1.0.0')
  s.add_development_dependency('shoulda', '~> 3.5.0')
  s.add_development_dependency "gem-release"
  
  s.require_paths = ["lib"]
end
