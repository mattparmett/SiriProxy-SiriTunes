require 'cora'
require 'siri_objects'
#require 'win32ole'

#######
# This is a "hello world" style plugin. It simply intercepts the phrase "text siri proxy" and responds
# with a message about the proxy being up and running (along with a couple other core features). This 
# is good base code for other plugins.
# 
# Remember to add other plugins to the "config.yml" file if you create them!
######

class SiriProxy::Plugin::SiriTunes < SiriProxy::Plugin
  def initialize(config)
    #if you have custom configuration options, process them here!
  end

  listen_for /itunes (.*)/i do |userAction|
	userAction = userAction.downcase
	itunes = WIN32OLE.new('iTunes.Application')
	if userAction == 'pause'
		itunes.PlayPause
		say "iTunes is now paused."
		request_completed
	elsif userAction == 'play'
		itunes.PlayPause
		say "iTunes is now playing."
		request_completed
	elsif userAction == 'next song'
		itunes.NextTrack
		say "Skipping to the next song."
		request_completed
	elsif userAction == 'previous song'
		itunes.PreviousTrack
		say "Skipping to the previous song."
		request_completed
	elsif userAction == 'lower the volume'
		itunes.SoundVolume = itunes.SoundVolume - 20
		say "Lowering the volume."
		request_completed
	elsif userAction == 'raise the volume'
		itunes.SoundVolume = itunes.SoundVolume + 20
		say "Raising the volume."
		request_completed
	end
  end
  listen_for /i tunes (.*)/i do |userAction|
	userAction = userAction.downcase
	itunes = WIN32OLE.new('iTunes.Application')
	if userAction == 'pause'
		itunes.PlayPause
		say "iTunes is now paused."
		request_completed
	elsif userAction == 'play'
		itunes.PlayPause
		say "iTunes is now playing."
		request_completed
	elsif userAction == 'next song'
		itunes.NextTrack
		say "Skipping to the next song."
		request_completed
	elsif userAction == 'previous song'
		itunes.PreviousTrack
		say "Skipping to the previous song."
		request_completed
	elsif userAction == 'lower the volume'
		itunes.SoundVolume = itunes.SoundVolume - 20
		say "Lowering the volume."
		request_completed
	elsif userAction == 'raise the volume'
		itunes.SoundVolume = itunes.SoundVolume + 20
		say "Raising the volume."
		request_completed
	end
  end
  
  listen_for /itunes put on (.*)/i do |name|
	name = name.split(' ').map {|w| w.capitalize }.join(' ')
	itunes = WIN32OLE.new('iTunes.Application')
	library = itunes.LibraryPlaylist
	tracks = library.Tracks
	song = tracks.ItemByName(songName)
	if !song
		#Song wasn't found.  Check if the user said an artist.
		artist_tracks = library.Search(songName, 2)
		if !artist_tracks #artist wasn't found. check if the user said an album.
			album_tracks = library.Search(songName, 3)
			if !album_tracks #album wasn't found.  We give up.
				say "Sorry, I couldn't find " + songName + " in your library."
				request_completed
			else
				songs = []
				for track in album_tracks
					songs << track
				end
				song = songs[0]
				if !song
					#We give up
					say "Sorry, I couldn't find " + songName + " in your library."
					request_completed
				else
					song.Play
					say "Playing " + songName + "."
					request_completed
				end
			end
		end
		songs = []
		for track in artist_tracks
			songs << track
		end
		song = songs[0]
		if !song
			#this should never happen. let's print a message if it does
			puts "The artist was found but the song is null."
			request_completed
		else
			songNum = rand(songs.length)
			song = songs[songNum]
			song.Play
			say "Playing " + songName + "."
			request_completed
		end
	else
		song.Play
		say "Playing " + songName + "."
		request_completed
	end
  end
end
