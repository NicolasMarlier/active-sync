# -*- encoding: utf-8 -*-

$:.push File.expand_path('../lib', __FILE__)

Gem::Specification.new do |s|
  s.name = "active-sync"
  s.summary = "A Ruby interface to Active sync."
  s.description = "A Ruby interface to Active sync"
  s.version = "0.1.0.4"
  s.platform = Gem::Platform::RUBY
  s.authors = ["Nicolas Marlier"]
  s.homepage = "http://github.com/nmarlier/active-sync"
  s.licenses = ['MIT']

  # runtime dependencies
  s.add_dependency "httpclient"
  
  # development dependencies
  s.add_development_dependency "rake"
  s.add_development_dependency "rspec", "~> 2.0"
  s.add_development_dependency "mocha", ">= 0.9"
  s.add_development_dependency "gem-release"
  
  s.require_paths = ["lib"]
end
