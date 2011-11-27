require 'cora'
require 'siri_objects'
require 'win32ole'

#######
# SiriTunes is a Siri Proxy plugin that allows you to play, pause, adjust the volume, skip to the next/previous track, and request specific songs, albums, and artists in iTunes on a Windows PC.

# Check the readme file for more detailed usage instructions

# SiriTunes was created by parm289.  You are free to use, modify, and redistribute this gem as long as proper credit is given to the original author.
######

class SiriProxy::Plugin::SiriTunes < SiriProxy::Plugin
  def initialize(config)
    #if you have custom configuration options, process them here!
  end

  def playReq(name)
	name = name.split(' ').map {|w| w.capitalize }.join(' ')
	itunes = WIN32OLE.new('iTunes.Application')
	library = itunes.LibraryPlaylist
	tracks = library.Tracks
	song = tracks.ItemByName(name)
	if !song
		#Song wasn't found.  Check if the user said an artist.
		artist_tracks = library.Search(name, 2)
		if !artist_tracks #artist wasn't found. check if the user said an album.
			album_tracks = library.Search(name, 3)
			if !album_tracks #album wasn't found.  We give up.
				return "Sorry, I couldn't find " + name + " in your library."
				request_completed
			else
				songs = []
				for track in album_tracks
					songs << track
				end
				song = songs[0]
				if !song
					#We give up
					return "Sorry, I couldn't find " + name + " in your library."
					request_completed
				else
					song.Play
					return "Playing " + name + "."
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
			return "Playing " + name + "."
			request_completed
		end
	else
		song.Play
		return "Playing " + name + "."
		request_completed
	end
  end
  
  def takeAction(userAction)
	userAction = userAction.downcase
	itunes = WIN32OLE.new('iTunes.Application')
	if userAction == 'pause' or userAction == 'pause '
		itunes.PlayPause
		return "iTunes is now paused."
		request_completed
	elsif userAction == 'play' or userAction == 'play '
		itunes.PlayPause
		return "iTunes is now playing."
		request_completed
	elsif userAction == 'next song' or userAction == 'next song '
		itunes.NextTrack
		return "Skipping to the next song."
		request_completed
	elsif userAction == 'previous song' or userAction == 'previous song '
		itunes.PreviousTrack
		return "Skipping to the previous song."
		request_completed
	elsif userAction == 'lower the volume' or userAction == 'lower the volume '
		itunes.SoundVolume = itunes.SoundVolume - 20
		return "Lowering the volume."
		request_completed
	elsif userAction == 'raise the volume' or userAction == 'raise the volume '
		itunes.SoundVolume = itunes.SoundVolume + 20
		return "Raising the volume."
		request_completed
	else
		return "Sorry, I didn't understand your request."
		request_completed
	end
  end
 
  listen_for /itunes put on (.*)/i do |name|
	response = playReq(name)
	say response
  end
  
  listen_for /i tunes put on (.*)/i do |name|
	response = playReq(name)
	say response
  end
  
  listen_for /itunes play (.*)/i do |name|
	if !name or name == "" or name == " "
		response = takeAction('play')
	else
		response = playReq(name)
	end
	say response
  end
  
  listen_for /i tunes play (.*)/i do |name|
	response = playReq(name)
	say response
  end
  
  listen_for /itunes (.*)/i do |userAction|
	response = takeAction(userAction)
	say response
  end
  
  listen_for /i tunes (.*)/i do |userAction|
	response = takeAction(userAction)
	say response
  end
end
